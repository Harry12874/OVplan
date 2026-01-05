const defaultColumns = [
  { header: "stop_name", field: "task.title" },
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
  { label: "Task title", value: "task.title" },
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

function resolveLocation(address = "") {
  const parts = address
    .split(",")
    .map((part) => part.trim())
    .filter(Boolean);
  if (parts.length >= 3) {
    return parts[parts.length - 2];
  }
  if (parts.length === 2) {
    return parts[1];
  }
  return address;
}

function normalizeOrderTaskType(taskType) {
  if (taskType === "pickup" || taskType === "packing" || taskType === "other") {
    return taskType;
  }
  return "delivery";
}

function buildTaskTitle({ kind, customer, order }) {
  const baseLabel = kind === "pickup" ? "PICKUP" : "DELIVERY";
  const name =
    kind === "pickup"
      ? order?.supplierName || order?.reference || customer?.storeName
      : customer?.storeName || order?.supplierName || order?.reference;
  const resolvedName = name || (order?.id ? `Order ${order.id}` : "Order");
  const address = customer?.fullAddress || "";
  const location = resolveLocation(address);
  if (!location) {
    return `${baseLabel}: ${resolvedName}`;
  }
  return `${baseLabel}: ${resolvedName} - ${location}`;
}

function buildSpokenitExportRecords({
  tasks,
  orders,
  customers,
  reps,
  start,
  end,
  repFilter,
  includeCompleted,
}) {
  const customerMap = new Map(customers.map((customer) => [customer.id, customer]));
  const orderMap = new Map(orders.map((order) => [order.id, order]));
  const repMap = new Map(reps.map((rep) => [rep.id, rep]));

  const deliveryTasks = tasks.filter((task) => {
    if (task.type !== "deliver") return false;
    if (!includeCompleted && task.status === "done") return false;
    if (task.dueDate < start || task.dueDate > end) return false;
    if (repFilter !== "all" && task.assignedRepId !== repFilter) return false;
    return true;
  });

  const deliveryRecords = deliveryTasks.map((task) => {
    const order = orderMap.get(task.orderId) || {};
    const customer = customerMap.get(task.customerId) || {};
    const rep = repMap.get(task.assignedRepId) || {};
    return {
      task: {
        ...task,
        type: "delivery",
        title: buildTaskTitle({ kind: "delivery", customer, order }),
      },
      order,
      customer,
      rep,
    };
  });

  const deliveryOrderIds = new Set(deliveryRecords.map((record) => record.order?.id).filter(Boolean));

  const pickupRecords = orders
    .filter((order) => normalizeOrderTaskType(order.taskType) === "pickup")
    .filter((order) => {
      if (!includeCompleted && order.status === "delivered") return false;
      if (order.deliveryDueDate < start || order.deliveryDueDate > end) return false;
      if (repFilter !== "all" && order.assignedRepId !== repFilter) return false;
      if (deliveryOrderIds.has(order.id)) return false;
      return true;
    })
    .map((order) => {
      const customer = customerMap.get(order.customerId) || {};
      const rep = repMap.get(order.assignedRepId) || {};
      return {
        task: {
          type: "pickup",
          dueDate: order.deliveryDueDate,
          assignedRepId: order.assignedRepId,
          title: buildTaskTitle({ kind: "pickup", customer, order }),
        },
        order,
        customer,
        rep,
      };
    });

  return [...deliveryRecords, ...pickupRecords];
}

export {
  defaultColumns,
  fieldOptions,
  buildCsv,
  downloadCsv,
  buildSpokenitExportRecords,
  buildTaskTitle,
  normalizeOrderTaskType,
};
