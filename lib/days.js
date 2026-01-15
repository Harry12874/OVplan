const MEL_TIMEZONE = "Australia/Melbourne";

const MON0_LABELS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];

const LABEL_TO_MON0 = {
  mon: 0,
  monday: 0,
  tue: 1,
  tues: 1,
  tuesday: 1,
  wed: 2,
  weds: 2,
  wednesday: 2,
  thu: 3,
  thur: 3,
  thurs: 3,
  thursday: 3,
  fri: 4,
  friday: 4,
  sat: 5,
  saturday: 5,
  sun: 6,
  sunday: 6,
};

function mon0ToLabel(index) {
  return MON0_LABELS[index] || "";
}

function labelsToMon0(label) {
  if (!label) return null;
  const normalized = String(label).trim().toLowerCase();
  if (!normalized) return null;
  return LABEL_TO_MON0[normalized] ?? null;
}

function getMelbourneDowMon0(date = new Date()) {
  const formatter = new Intl.DateTimeFormat("en-AU", {
    timeZone: MEL_TIMEZONE,
    weekday: "short",
  });
  const label = formatter.format(date);
  const mapped = labelsToMon0(label);
  if (mapped !== null && mapped !== undefined) {
    return mapped;
  }
  const sun0 = date.getDay();
  return (sun0 + 6) % 7;
}

export { MEL_TIMEZONE, getMelbourneDowMon0, mon0ToLabel, labelsToMon0 };
