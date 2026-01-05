const defaultColumns = [
  { header: "stop_name", field: "customer.storeName" },
  { header: "address", field: "customer.fullAddress" },
  { header: "notes", field: "customer.deliveryNotes" },
  { header: "phone", field: "customer.phone" },
  { header: "email", field: "customer.email" },
  { header: "rep_name", field: "rep.name" },
  { header: "due_date", field: "task.dueDate" },
  { header: "order_id", field: "order.id" },
  { header: "customer_id", field: "customer.id" },
  { header: "service_time_minutes", field: "order.serviceTime" },
];

const fieldOptions = [
  { label: "Store name", value: "customer.storeName" },
  { label: "Address", value: "customer.fullAddress" },
  { label: "Delivery notes", value: "customer.deliveryNotes" },
  { label: "Customer phone", value: "customer.phone" },
  { label: "Customer email", value: "customer.email" },
  { label: "Rep name", value: "rep.name" },
  { label: "Delivery due date", value: "task.dueDate" },
  { label: "Order id", value: "order.id" },
  { label: "Customer id", value: "customer.id" },
  { label: "Order notes", value: "order.internalNotes" },
  { label: "Order channel", value: "order.channel" },
  { label: "Service time (minutes)", value: "order.serviceTime" },
];

function resolveField(record, path) {
  if (!path) return "";
  const parts = path.split(".");
  return parts.reduce((acc, key) => (acc ? acc[key] : ""), record) ?? "";
}

function toCsvValue(value) {
  if (value === null || value === undefined) return "";
  const stringValue = String(value);
  if (stringValue.includes(",") || stringValue.includes("\n") || stringValue.includes('"')) {
    return `"${stringValue.replace(/"/g, '""')}"`;
  }
  return stringValue;
}

function buildCsv(records, columns) {
  const headerRow = columns.map((col) => toCsvValue(col.header)).join(",");
  const rows = records.map((record) => {
    return columns
      .map((col) => toCsvValue(resolveField(record, col.field)))
      .join(",");
  });
  return [headerRow, ...rows].join("\n");
}

function downloadCsv(filename, csvString) {
  const blob = new Blob([csvString], { type: "text/csv;charset=utf-8;" });
  const link = document.createElement("a");
  const url = URL.createObjectURL(blob);
  link.setAttribute("href", url);
  link.setAttribute("download", filename);
  link.style.visibility = "hidden";
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

export { defaultColumns, fieldOptions, buildCsv, downloadCsv };
