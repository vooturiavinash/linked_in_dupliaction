import * as XLSX from 'xlsx';

export type Lead = {
  row: number;
  date: string;
  contactName: string;
  designation: string;
  school: string;
  location: string;
  phone1: string;
  phone2: string;
  phone3: string;
  email1: string;
  email2: string;
  ownerName: string;
  linkedinUrl: string;
  normalizedUrl: string | null;
};

export type LeadStatus = 'duplicate' | 'unique' | 'missing' | 'invalid';

export type CheckedLead = Lead & {
  status: LeadStatus;
  duplicateCount: number;
  groupId: number | null;
};

export type DuplicateGroup = {
  id: number;
  key: string;
  url: string;
  count: number;
  rowNumbers: number[];
  names: string[];
};

export type Stats = {
  total: number;
  withUrl: number;
  missing: number;
  invalid: number;
  distinct: number;
  duplicateGroups: number;
  duplicateRows: number;
};

export type Analysis = {
  leads: CheckedLead[];
  groups: DuplicateGroup[];
  stats: Stats;
};

export type Snapshot = {
  ok: boolean;
  error?: string;
  fileName: string;
  filePath: string;
  modifiedAt: string | null;
  leads: Lead[];
};

export type LeadField = Exclude<keyof Lead, 'row' | 'normalizedUrl'>;

export const EDITABLE_FIELDS: LeadField[] = [
  'date',
  'contactName',
  'designation',
  'school',
  'location',
  'phone1',
  'phone2',
  'phone3',
  'email1',
  'email2',
  'ownerName',
  'linkedinUrl',
];

export type RowIdentity = {
  contactName: string;
  linkedinUrl: string;
  school: string;
};

export type SheetLayout = {
  sheetName: string;
  headerExcelRow: number;
  columns: Array<{ index: number; field: LeadField }>;
};

const HEADER_FIELDS: Record<string, LeadField> = {
  date: 'date',
  'contact name': 'contactName',
  name: 'contactName',
  designation: 'designation',
  title: 'designation',
  school: 'school',
  company: 'school',
  organisation: 'school',
  organization: 'school',
  location: 'location',
  city: 'location',
  'phone 1': 'phone1',
  phone1: 'phone1',
  phone: 'phone1',
  mobile: 'phone1',
  'phone 2': 'phone2',
  phone2: 'phone2',
  'phone 3': 'phone3',
  phone3: 'phone3',
  'email 1': 'email1',
  email1: 'email1',
  email: 'email1',
  'email 2': 'email2',
  email2: 'email2',
  'owner name': 'ownerName',
  owner: 'ownerName',
  'linkedin url': 'linkedinUrl',
  linkedin: 'linkedinUrl',
  'linkedin link': 'linkedinUrl',
  'profile url': 'linkedinUrl',
};

const PLACEHOLDER = /^(n\/a|na|none|nil|null|-+|—|tbd|not available|no|pending)?$/i;

const GROUP_COLORS = [
  '#d4382c',
  '#c2410c',
  '#a16207',
  '#7c3aed',
  '#0369a1',
  '#0f766e',
  '#be185d',
  '#4338ca',
];

export function groupColor(groupId: number): string {
  return GROUP_COLORS[(groupId - 1) % GROUP_COLORS.length];
}

function normHeader(value: string): string {
  return value.replace(/[\r\n]+/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase();
}

function fieldFor(header: string): LeadField | null {
  const key = normHeader(header);
  if (!key) return null;
  if (key.includes('linkedin') || key.includes('lnkd.in')) return 'linkedinUrl';
  return HEADER_FIELDS[key] ?? null;
}

function clean(value: string): string {
  return value.replace(/[\r\n]+/g, ' ').replace(/[ \t]+/g, ' ').trim();
}

function numberToText(value: number): string {
  if (!Number.isFinite(value)) return '';
  if (Number.isInteger(value)) return String(value);
  return String(value);
}

function excelSerialToDate(serial: number): string {
  const utc = new Date(Date.UTC(1899, 11, 30) + Math.round(serial) * 86400000);
  const day = String(utc.getUTCDate()).padStart(2, '0');
  const month = String(utc.getUTCMonth() + 1).padStart(2, '0');
  return `${day}/${month}/${utc.getUTCFullYear()}`;
}

function asDate(value: string | number | null): string {
  if (value == null || value === '') return '';
  if (typeof value === 'number' && value > 20000 && value < 80000) {
    return excelSerialToDate(value);
  }
  const text = clean(String(value));
  if (/^\d{5}(\.0+)?$/.test(text)) {
    const serial = Number(text);
    if (serial > 20000 && serial < 80000) return excelSerialToDate(serial);
  }
  return text;
}

function asText(value: string | number | null): string {
  if (value == null || value === '') return '';
  if (typeof value === 'number') return numberToText(value);
  return clean(value);
}

function readRaw(sheet: XLSX.WorkSheet, row: number, column: number): string | number | null {
  const cell = sheet[XLSX.utils.encode_cell({ r: row, c: column })] as XLSX.CellObject | undefined;
  if (!cell || cell.v == null || cell.v === '') return null;
  if (typeof cell.v === 'number') return cell.v;
  if (cell.v instanceof Date) {
    const day = String(cell.v.getDate()).padStart(2, '0');
    const month = String(cell.v.getMonth() + 1).padStart(2, '0');
    return `${day}/${month}/${cell.v.getFullYear()}`;
  }
  return clean(String(cell.v));
}

function safeDecode(value: string): string {
  try {
    return decodeURIComponent(value);
  } catch {
    return value;
  }
}

export function normalizeLinkedIn(input: string): string | null {
  const raw = clean(input);
  if (!raw || PLACEHOLDER.test(raw)) return null;
  if (!/linkedin\.com|lnkd\.in/i.test(raw)) return null;

  let candidate = raw;
  if (!/^https?:\/\//i.test(candidate)) {
    candidate = `https://${candidate.replace(/^\/\//, '')}`;
  }

  try {
    const url = new URL(candidate);
    const host = url.hostname.toLowerCase().replace(/^www\./, '');
    if (host === 'lnkd.in') {
      const short = url.pathname.replace(/\/+$/, '').toLowerCase();
      return `lnkd.in${short}`;
    }
    if (!host.endsWith('linkedin.com')) return null;

    const path = safeDecode(url.pathname).replace(/\/+/g, '/').replace(/\/+$/, '').toLowerCase();
    const person = path.match(/\/in\/([^/]+)/);
    if (person) return `${host}/in/${person[1]}`;
    const legacy = path.match(/\/pub\/([^/]+)/);
    if (legacy) return `${host}/in/${legacy[1]}`;
    const company = path.match(/\/company\/([^/]+)/);
    if (company) return `${host}/company/${company[1]}`;
    return `${host}${path}`;
  } catch {
    return raw.toLowerCase().replace(/\/+$/, '').split('?')[0].replace(/^https?:\/\//, '').replace(/^www\./, '');
  }
}

type HeaderHit = {
  row: number;
  score: number;
  columns: Array<{ index: number; field: LeadField }>;
};

function findHeader(sheet: XLSX.WorkSheet): HeaderHit | null {
  const ref = sheet['!ref'];
  if (!ref) return null;
  const range = XLSX.utils.decode_range(ref);
  const lastRow = Math.min(range.e.r, range.s.r + 20);
  let best: HeaderHit | null = null;

  for (let row = range.s.r; row <= lastRow; row += 1) {
    const columns: Array<{ index: number; field: LeadField }> = [];
    const used = new Set<LeadField>();
    let score = 0;
    for (let column = range.s.c; column <= range.e.c; column += 1) {
      const raw = readRaw(sheet, row, column);
      const field = fieldFor(raw == null ? '' : String(raw));
      if (!field || used.has(field)) continue;
      used.add(field);
      columns.push({ index: column, field });
      score += field === 'linkedinUrl' ? 4 : 1;
    }
    if (score >= 3 && (!best || score > best.score)) {
      best = { row, score, columns };
    }
  }

  return best;
}

function blankLead(row: number): Lead {
  return {
    row,
    date: '',
    contactName: '',
    designation: '',
    school: '',
    location: '',
    phone1: '',
    phone2: '',
    phone3: '',
    email1: '',
    email2: '',
    ownerName: '',
    linkedinUrl: '',
    normalizedUrl: null,
  };
}

function readRows(sheet: XLSX.WorkSheet, header: HeaderHit): Lead[] {
  const ref = sheet['!ref'];
  if (!ref) return [];
  const range = XLSX.utils.decode_range(ref);
  const leads: Lead[] = [];

  for (let row = header.row + 1; row <= range.e.r; row += 1) {
    const lead = blankLead(row + 1);
    let any = false;
    for (const column of header.columns) {
      const raw = readRaw(sheet, row, column.index);
      const text = column.field === 'date' ? asDate(raw) : asText(raw);
      if (!text) continue;
      lead[column.field] = text;
      any = true;
    }
    if (!any) continue;
    lead.normalizedUrl = normalizeLinkedIn(lead.linkedinUrl);
    leads.push(lead);
  }

  return leads;
}

export function inspectWorkbook(data: ArrayBuffer | Uint8Array): { leads: Lead[]; layout: SheetLayout | null } {
  const bytes = data instanceof ArrayBuffer ? new Uint8Array(data) : new Uint8Array(data);
  const workbook = XLSX.read(bytes, { type: 'array', cellDates: false });
  let bestSheet: XLSX.WorkSheet | null = null;
  let bestHeader: HeaderHit | null = null;
  let bestName = '';

  for (const name of workbook.SheetNames) {
    const sheet = workbook.Sheets[name];
    const header = findHeader(sheet);
    if (header && (!bestHeader || header.score > bestHeader.score)) {
      bestHeader = header;
      bestSheet = sheet;
      bestName = name;
    }
  }

  if (!bestSheet || !bestHeader) return { leads: [], layout: null };
  return {
    leads: readRows(bestSheet, bestHeader),
    layout: {
      sheetName: bestName,
      headerExcelRow: bestHeader.row + 1,
      columns: bestHeader.columns,
    },
  };
}

export function parseSheet(data: ArrayBuffer | Uint8Array): Lead[] {
  return inspectWorkbook(data).leads;
}

export function analyze(leads: Lead[]): Analysis {
  const buckets = new Map<string, Lead[]>();
  for (const lead of leads) {
    if (!lead.normalizedUrl) continue;
    const list = buckets.get(lead.normalizedUrl) ?? [];
    list.push(lead);
    buckets.set(lead.normalizedUrl, list);
  }

  const groups: DuplicateGroup[] = [];
  const groupByKey = new Map<string, DuplicateGroup>();
  for (const [key, list] of buckets) {
    if (list.length < 2) continue;
    const group: DuplicateGroup = {
      id: groups.length + 1,
      key,
      url: list[0].linkedinUrl,
      count: list.length,
      rowNumbers: list.map((lead) => lead.row),
      names: list.map((lead) => lead.contactName || 'Unnamed'),
    };
    groups.push(group);
    groupByKey.set(key, group);
  }

  const checked: CheckedLead[] = leads.map((lead) => {
    const hasText = clean(lead.linkedinUrl).length > 0 && !PLACEHOLDER.test(clean(lead.linkedinUrl));
    if (!lead.normalizedUrl) {
      return {
        ...lead,
        status: hasText ? 'invalid' : 'missing',
        duplicateCount: 0,
        groupId: null,
      };
    }
    const count = buckets.get(lead.normalizedUrl)?.length ?? 1;
    const group = groupByKey.get(lead.normalizedUrl);
    if (count >= 2 && group) {
      return { ...lead, status: 'duplicate', duplicateCount: count, groupId: group.id };
    }
    return { ...lead, status: 'unique', duplicateCount: 1, groupId: null };
  });

  const stats: Stats = {
    total: checked.length,
    withUrl: checked.filter((lead) => lead.status === 'duplicate' || lead.status === 'unique').length,
    missing: checked.filter((lead) => lead.status === 'missing').length,
    invalid: checked.filter((lead) => lead.status === 'invalid').length,
    distinct: buckets.size,
    duplicateGroups: groups.length,
    duplicateRows: checked.filter((lead) => lead.status === 'duplicate').length,
  };

  return { leads: checked, groups, stats };
}
