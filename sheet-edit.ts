import ExcelJS from 'exceljs';
import { inspectWorkbook, normalizeLinkedIn, type Lead, type LeadField, type RowIdentity, EDITABLE_FIELDS } from './leads';

export class SheetConflict extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'SheetConflict';
  }
}

function clean(value: string): string {
  return value.replace(/[\r\n]+/g, ' ').replace(/[ \t]+/g, ' ').trim();
}

function sameText(left: string, right: string): boolean {
  return clean(left).toLowerCase() === clean(right).toLowerCase();
}

function matchesIdentity(lead: Lead, expect: RowIdentity): boolean {
  const expectedUrl = normalizeLinkedIn(expect.linkedinUrl);
  const hasName = clean(expect.contactName).length > 0 || clean(lead.contactName).length > 0;
  const hasUrl = Boolean(expectedUrl || lead.normalizedUrl);
  if (hasName && !sameText(lead.contactName, expect.contactName)) return false;
  if (hasUrl && lead.normalizedUrl !== expectedUrl) return false;
  if (!hasName && !hasUrl) return sameText(lead.school, expect.school);
  return true;
}

export function findLiveLead(leads: Lead[], row: number, expect: RowIdentity, headerExcelRow: number): Lead {
  if (!Number.isInteger(row) || row <= headerExcelRow) {
    throw new SheetConflict('That row cannot be changed.');
  }
  const lead = leads.find((item) => item.row === row);
  if (!lead) {
    throw new SheetConflict('That row is no longer in the spreadsheet. It was removed or the file changed.');
  }
  if (!matchesIdentity(lead, expect)) {
    const who = lead.contactName || lead.school || 'a different lead';
    throw new SheetConflict(`The spreadsheet changed. Row ${row} is now ${who}. Check again before saving.`);
  }
  return lead;
}

async function loadWorkbook(bytes: Uint8Array): Promise<ExcelJS.Workbook> {
  const workbook = new ExcelJS.Workbook();
  const data = Buffer.from(bytes) as unknown as Parameters<ExcelJS.Xlsx['load']>[0];
  await workbook.xlsx.load(data);
  return workbook;
}

export async function applyLeadEdit(
  bytes: Uint8Array,
  row: number,
  expect: RowIdentity,
  fields: Partial<Record<LeadField, string>>,
): Promise<Buffer> {
  const { leads, layout } = inspectWorkbook(bytes);
  if (!layout) throw new SheetConflict('Could not find the lead columns in the spreadsheet.');
  const current = findLiveLead(leads, row, expect, layout.headerExcelRow);
  const workbook = await loadWorkbook(bytes);
  const sheet = workbook.getWorksheet(layout.sheetName);
  if (!sheet) throw new SheetConflict('Could not find the worksheet.');

  const excelRow = sheet.getRow(row);
  let changed = false;
  for (const column of layout.columns) {
    if (!EDITABLE_FIELDS.includes(column.field) || !(column.field in fields)) continue;
    const next = clean(fields[column.field] ?? '');
    if (next === current[column.field]) continue;
    excelRow.getCell(column.index + 1).value = next || null;
    changed = true;
  }

  if (!changed) return Buffer.from(bytes);
  excelRow.commit();
  return Buffer.from(await workbook.xlsx.writeBuffer());
}

export async function applyLeadDelete(bytes: Uint8Array, row: number, expect: RowIdentity): Promise<Buffer> {
  const { leads, layout } = inspectWorkbook(bytes);
  if (!layout) throw new SheetConflict('Could not find the lead columns in the spreadsheet.');
  findLiveLead(leads, row, expect, layout.headerExcelRow);
  const workbook = await loadWorkbook(bytes);
  const sheet = workbook.getWorksheet(layout.sheetName);
  if (!sheet) throw new SheetConflict('Could not find the worksheet.');
  sheet.spliceRows(row, 1);
  return Buffer.from(await workbook.xlsx.writeBuffer());
}
