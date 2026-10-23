import fs from 'node:fs';
import http from 'node:http';
import path from 'node:path';
import { applyLeadDelete, applyLeadEdit, SheetConflict } from '../src/sheet-edit';
import { EDITABLE_FIELDS, parseSheet, type LeadField, type RowIdentity, type Snapshot } from '../src/leads';

const FILE = process.env.LDC_XLSX ?? 'C:\\Users\\User\\Downloads\\Linkedin Leads updated.xlsx';
const PORT = Number(process.env.LDC_PORT ?? 8787);

const clients = new Set<http.ServerResponse>();
let notifyTimer: NodeJS.Timeout | null = null;
let seenMtime = 0;
let muteUntil = 0;
let writeQueue: Promise<unknown> = Promise.resolve();

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => {
    setTimeout(resolve, ms);
  });
}

function errorMessage(error: unknown): string {
  const code = typeof error === 'object' && error && 'code' in error ? String(error.code) : '';
  if (code === 'EBUSY' || code === 'EPERM' || code === 'EACCES') {
    return 'The spreadsheet is locked, usually because it is open in Excel. Save it and LDC will check again.';
  }
  if (code === 'ENOENT') {
    return `Could not find the spreadsheet at ${FILE}.`;
  }
  return error instanceof Error ? error.message : 'Could not read the spreadsheet.';
}

async function readSnapshot(): Promise<Snapshot> {
  const fileName = path.basename(FILE);
  let lastError: unknown;

  for (let attempt = 0; attempt < 5; attempt += 1) {
    try {
      const bytes = fs.readFileSync(FILE);
      if (bytes.length < 32) throw new Error('The spreadsheet is still being saved.');
      const leads = parseSheet(bytes);
      const stat = fs.statSync(FILE);
      return {
        ok: true,
        fileName,
        filePath: FILE,
        modifiedAt: stat.mtime.toISOString(),
        leads,
      };
    } catch (error) {
      lastError = error;
      await delay(220);
    }
  }

  return {
    ok: false,
    error: errorMessage(lastError),
    fileName,
    filePath: FILE,
    modifiedAt: null,
    leads: [],
  };
}

function notifyClients(): void {
  if (Date.now() < muteUntil) return;
  if (notifyTimer) clearTimeout(notifyTimer);
  notifyTimer = setTimeout(() => {
    if (Date.now() < muteUntil) return;
    const payload = `data: ${JSON.stringify({ type: 'change', at: new Date().toISOString() })}\n\n`;
    for (const client of clients) client.write(payload);
  }, 700);
}

function enqueue<T>(job: () => Promise<T>): Promise<T> {
  const run = writeQueue.then(job, job);
  writeQueue = run.then(
    () => undefined,
    () => undefined,
  );
  return run;
}

function readBody(req: http.IncomingMessage): Promise<unknown> {
  return new Promise((resolve, reject) => {
    const chunks: Buffer[] = [];
    let size = 0;
    req.on('data', (chunk: Buffer) => {
      size += chunk.length;
      if (size > 1_000_000) {
        reject(new Error('The edit was too large.'));
        req.destroy();
        return;
      }
      chunks.push(chunk);
    });
    req.on('end', () => {
      try {
        resolve(JSON.parse(Buffer.concat(chunks).toString('utf8') || '{}') as unknown);
      } catch {
        reject(new Error('The edit could not be read.'));
      }
    });
    req.on('error', reject);
  });
}

function asIdentity(value: unknown): RowIdentity | null {
  if (!value || typeof value !== 'object') return null;
  const record = value as Record<string, unknown>;
  if (typeof record.contactName !== 'string' || typeof record.linkedinUrl !== 'string' || typeof record.school !== 'string') {
    return null;
  }
  return { contactName: record.contactName, linkedinUrl: record.linkedinUrl, school: record.school };
}

function asFields(value: unknown): Partial<Record<LeadField, string>> | null {
  if (!value || typeof value !== 'object') return null;
  const record = value as Record<string, unknown>;
  const fields: Partial<Record<LeadField, string>> = {};
  for (const field of EDITABLE_FIELDS) {
    if (!(field in record)) continue;
    if (typeof record[field] !== 'string') return null;
    fields[field] = record[field];
  }
  return fields;
}

function replaceFile(bytes: Buffer): void {
  const temporary = `${FILE}.ldc-writing`;
  const backup = `${FILE}.ldc-backup`;
  fs.writeFileSync(temporary, bytes);
  try {
    if (fs.existsSync(backup)) fs.unlinkSync(backup);
    fs.renameSync(FILE, backup);
  } catch (error) {
    try {
      fs.unlinkSync(temporary);
    } catch {
      /* The new copy could not be kept. */
    }
    throw error;
  }
  try {
    fs.renameSync(temporary, FILE);
  } catch (error) {
    try {
      fs.renameSync(backup, FILE);
    } catch {
      /* The backup still has the previous spreadsheet. */
    }
    throw error;
  }
  try {
    fs.unlinkSync(backup);
  } catch {
    /* A later save removes the backup. */
  }
}

async function writeSpreadsheet(build: (bytes: Uint8Array) => Promise<Buffer>): Promise<Snapshot> {
  return enqueue(async () => {
    const current = fs.readFileSync(FILE);
    let next: Buffer;
    try {
      next = await build(current);
    } catch (error) {
      if (error instanceof SheetConflict) {
        const snapshot = await readSnapshot();
        return { ...snapshot, ok: false, error: error.message };
      }
      throw error;
    }
    if (Buffer.compare(current, next) === 0) return readSnapshot();
    muteUntil = Date.now() + 5000;
    try {
      replaceFile(next);
      const stat = fs.statSync(FILE);
      seenMtime = stat.mtimeMs;
    } catch (error) {
      muteUntil = 0;
      throw error;
    }
    return readSnapshot();
  });
}

function watchSpreadsheet(): void {
  fs.watchFile(FILE, { interval: 1000 }, (current, previous) => {
    if (current.mtimeMs !== previous.mtimeMs) notifyClients();
  });

  setInterval(() => {
    fs.stat(FILE, (error, stat) => {
      if (error || !stat) return;
      if (seenMtime && stat.mtimeMs !== seenMtime) notifyClients();
      seenMtime = stat.mtimeMs;
    });
  }, 1500);
}

function sendJson(res: http.ServerResponse, status: number, body: unknown): void {
  const payload = JSON.stringify(body);
  res.writeHead(status, {
    'Content-Type': 'application/json; charset=utf-8',
    'Cache-Control': 'no-store',
    'Content-Length': Buffer.byteLength(payload),
  });
  res.end(payload);
}

const server = http.createServer((req, res) => {
  const url = new URL(req.url ?? '/', 'http://127.0.0.1');

  if (req.method === 'GET' && url.pathname === '/api/snapshot') {
    void readSnapshot()
      .then((snapshot) => sendJson(res, snapshot.ok ? 200 : 200, snapshot))
      .catch((error: unknown) => {
        sendJson(res, 500, {
          ok: false,
          error: errorMessage(error),
          fileName: path.basename(FILE),
          filePath: FILE,
          modifiedAt: null,
          leads: [],
        } satisfies Snapshot);
      });
    return;
  }

  if (req.method === 'POST' && (url.pathname === '/api/leads/update' || url.pathname === '/api/leads/delete')) {
    const removing = url.pathname === '/api/leads/delete';
    void readBody(req)
      .then(async (body) => {
        const record = body && typeof body === 'object' ? (body as Record<string, unknown>) : {};
        const row = Number(record.row);
        const expect = asIdentity(record.expect);
        if (!expect || !Number.isInteger(row)) {
          sendJson(res, 400, {
            ok: false,
            error: 'The row to change was missing.',
            fileName: path.basename(FILE),
            filePath: FILE,
            modifiedAt: null,
            leads: [],
          } satisfies Snapshot);
          return;
        }
        const snapshot = await writeSpreadsheet((bytes) => {
          if (removing) return applyLeadDelete(bytes, row, expect);
          const fields = asFields(record.fields);
          if (!fields) throw new SheetConflict('The edited fields could not be read.');
          return applyLeadEdit(bytes, row, expect, fields);
        });
        sendJson(res, 200, snapshot);
      })
      .catch((error: unknown) => {
        sendJson(res, 200, {
          ok: false,
          error: errorMessage(error),
          fileName: path.basename(FILE),
          filePath: FILE,
          modifiedAt: null,
          leads: [],
        } satisfies Snapshot);
      });
    return;
  }

  if (req.method === 'GET' && url.pathname === '/api/events') {
    res.writeHead(200, {
      'Content-Type': 'text/event-stream',
      'Cache-Control': 'no-cache, no-transform',
      Connection: 'keep-alive',
    });
    res.write('retry: 2000\n\n');
    clients.add(res);
    const ping = setInterval(() => res.write(': ping\n\n'), 15000);
    req.on('close', () => {
      clearInterval(ping);
      clients.delete(res);
    });
    return;
  }

  sendJson(res, 404, { ok: false, error: 'Not found' });
});

server.listen(PORT, '127.0.0.1', () => {
  watchSpreadsheet();
  console.log(`LDC is watching ${FILE}`);
  console.log(`API ready at http://127.0.0.1:${PORT}`);
});

process.on('SIGINT', () => {
  fs.unwatchFile(FILE);
  process.exit(0);
});
