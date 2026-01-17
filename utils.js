import { MEL_TIMEZONE, getMelbourneDowMon0 } from "./lib/days.js";

const channelLabels = {
  portal: "Portal",
  sms: "SMS",
  email: "Email",
  phone: "Phone",
  rep_in_person: "Sales rep",
};

function uuid() {
  return crypto.randomUUID();
}

function todayKey() {
  return dateKeyFromDate(new Date());
}

function dateKeyFromDate(date) {
  const parts = new Intl.DateTimeFormat("en-AU", {
    timeZone: MEL_TIMEZONE,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).formatToParts(date);
  const lookup = Object.fromEntries(parts.map((p) => [p.type, p.value]));
  return `${lookup.year}-${lookup.month}-${lookup.day}`;
}

function dateKeyToDate(dateKey) {
  if (!dateKey) return new Date();
  const [year, month, day] = dateKey.split("-").map(Number);
  return new Date(Date.UTC(year, month - 1, day, 12));
}

function addDays(dateKey, days) {
  const base = dateKeyToDate(dateKey);
  base.setUTCDate(base.getUTCDate() + days);
  return dateKeyFromDate(base);
}

function formatDate(dateKey) {
  if (!dateKey) return "";
  return new Intl.DateTimeFormat("en-AU", {
    timeZone: MEL_TIMEZONE,
    weekday: "short",
    day: "2-digit",
    month: "short",
  }).format(dateKeyToDate(dateKey));
}

function formatDateTime(isoString) {
  if (!isoString) return "";
  return new Intl.DateTimeFormat("en-AU", {
    timeZone: MEL_TIMEZONE,
    dateStyle: "medium",
    timeStyle: "short",
  }).format(new Date(isoString));
}

function parseSuburb(address = "") {
  const parts = address.split(",").map((part) => part.trim());
  if (parts.length >= 2) {
    return parts[parts.length - 2];
  }
  return "";
}

function normalizeSearch(value) {
  return value.toLowerCase().trim();
}

function matchesSearch(value, term) {
  if (!term) return true;
  return normalizeSearch(value).includes(normalizeSearch(term));
}

function channelLabel(channel) {
  return channelLabels[channel] || channel;
}

function toISO(date = new Date()) {
  return date.toISOString();
}

function startOfWeek(dateKey) {
  const date = dateKeyToDate(dateKey);
  const day = dayIndexFromDateKey(dateKey);
  const diff = -day;
  date.setUTCDate(date.getUTCDate() + diff);
  return dateKeyFromDate(date);
}

function dateRange(startKey, days) {
  const keys = [];
  for (let i = 0; i < days; i += 1) {
    keys.push(addDays(startKey, i));
  }
  return keys;
}

function normalizeHeader(header = "") {
  return header
    .replace(/^\uFEFF/, "")
    .trim()
    .replace(/\*/g, "")
    .toLowerCase()
    .replace(/[\s_-]+/g, "");
}

function parseCsv(text) {
  const rows = [];
  let current = "";
  let row = [];
  let inQuotes = false;

  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];
    const next = text[i + 1];

    if (char === "\"" && inQuotes && next === "\"") {
      current += "\"";
      i += 1;
      continue;
    }

    if (char === "\"") {
      inQuotes = !inQuotes;
      continue;
    }

    if (char === "," && !inQuotes) {
      row.push(current);
      current = "";
      continue;
    }

    if ((char === "\n" || char === "\r") && !inQuotes) {
      if (char === "\r" && next === "\n") {
        i += 1;
      }
      row.push(current);
      if (row.length > 1 || row[0] !== "") {
        rows.push(row);
      }
      row = [];
      current = "";
      continue;
    }

    current += char;
  }

  if (current.length || row.length) {
    row.push(current);
    rows.push(row);
  }

  return rows;
}

function detectNumericColumns(headers, rows) {
  const numericIndices = new Set();
  rows.forEach((row) => {
    row.forEach((value, index) => {
      const trimmed = String(value || "").trim();
      if (trimmed && !Number.isNaN(Number(trimmed))) {
        numericIndices.add(index);
      }
    });
  });
  return headers
    .map((header, index) => ({ header, index }))
    .filter((item) => numericIndices.has(item.index));
}

const ISO_DAY_LABELS = {
  mon: 1,
  monday: 1,
  tue: 2,
  tues: 2,
  tuesday: 2,
  wed: 3,
  weds: 3,
  wednesday: 3,
  thu: 4,
  thur: 4,
  thurs: 4,
  thursday: 4,
  fri: 5,
  friday: 5,
  sat: 6,
  saturday: 6,
  sun: 7,
  sunday: 7,
};

function normaliseDays(input) {
  if (input === null || input === undefined) return [];
  let values = [];
  if (Array.isArray(input)) {
    values = input;
  } else if (typeof input === "string") {
    const trimmed = input.trim();
    if (!trimmed) return [];
    if (trimmed.startsWith("[") && trimmed.endsWith("]")) {
      try {
        const parsed = JSON.parse(trimmed);
        values = Array.isArray(parsed) ? parsed : [trimmed];
      } catch (error) {
        values = trimmed.split(/[,\s]+/);
      }
    } else {
      values = trimmed.split(/[,\s]+/);
    }
  } else {
    values = [input];
  }

  const numeric = [];
  const labels = [];
  values.forEach((value) => {
    if (value === null || value === undefined || value === "") return;
    if (typeof value === "number" && Number.isFinite(value)) {
      numeric.push(Math.trunc(value));
      return;
    }
    const stringValue = String(value).trim();
    if (!stringValue) return;
    const asNumber = Number(stringValue);
    if (!Number.isNaN(asNumber)) {
      numeric.push(Math.trunc(asNumber));
      return;
    }
    labels.push(stringValue.toLowerCase());
  });

  const isoFromLabels = labels
    .map((label) => ISO_DAY_LABELS[label])
    .filter((value) => value !== undefined);

  let isoFromNumbers = [];
  if (numeric.length) {
    const isSun0 = numeric.every((value) => value >= 0 && value <= 6);
    isoFromNumbers = numeric
      .map((value) => (isSun0 ? (value === 0 ? 7 : value) : value))
      .filter((value) => value >= 1 && value <= 7);
  }

  const combined = isoFromLabels.concat(isoFromNumbers).filter((value) => value >= 1 && value <= 5);
  return Array.from(new Set(combined)).sort((a, b) => a - b);
}

function dayIndexFromDateKey(dateKey) {
  if (!dateKey) return 0;
  return getMelbourneDowMon0(dateKeyToDate(dateKey));
}

export {
  MEL_TIMEZONE,
  uuid,
  todayKey,
  dateKeyFromDate,
  dateKeyToDate,
  dayIndexFromDateKey,
  addDays,
  formatDate,
  formatDateTime,
  parseSuburb,
  normalizeSearch,
  matchesSearch,
  channelLabel,
  toISO,
  startOfWeek,
  dateRange,
  normalizeHeader,
  parseCsv,
  detectNumericColumns,
  normaliseDays,
};
