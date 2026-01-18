import assert from "node:assert/strict";
import { buildSpokenitExportRecords } from "../export.js";

const baseCustomers = [
  {
    id: "cust-1",
    storeName: "Fresh Market",
    fullAddress: "12 Main St, Richmond VIC 3121",
  },
];

const baseReps = [{ id: "rep-1", name: "Jamie Cole" }];

function runExport({
  tasks,
  orders,
  start = "2025-01-01",
  end = "2025-01-02",
  repFilter = "all",
  includeCompleted = true,
}) {
  return buildSpokenitExportRecords({
    tasks,
    orders,
    customers: baseCustomers,
    reps: baseReps,
    start,
    end,
    repFilter,
    includeCompleted,
  });
}

{
  const records = runExport({
    tasks: [
      {
        id: "task-1",
        type: "deliver",
        orderId: "order-1",
        customerId: "cust-1",
        dueDate: "2025-01-02",
        status: "todo",
        assignedRepId: "rep-1",
      },
      {
        id: "task-pack",
        type: "pack",
        orderId: "order-1",
        customerId: "cust-1",
        dueDate: "2025-01-01",
        status: "todo",
        assignedRepId: "rep-1",
      },
    ],
    orders: [
      {
        id: "order-1",
        customerId: "cust-1",
        assignedRepId: "rep-1",
        deliveryDueDate: "2025-01-02",
        status: "received",
      },
    ],
  });

  assert.equal(records.length, 2);
  assert.ok(records.some((record) => record.task.eventType === "delivery"));
  assert.ok(records.some((record) => record.task.eventType === "order"));
  assert.ok(records.some((record) => record.task.title === "DELIVERY: Fresh Market - Richmond VIC 3121"));
  assert.ok(records.some((record) => record.task.title === "ORDER: Fresh Market - Richmond VIC 3121"));
}

{
  const records = runExport({
    tasks: [],
    orders: [
      {
        id: "order-pickup",
        customerId: "cust-1",
        assignedRepId: "rep-1",
        deliveryDueDate: "2025-01-02",
        status: "received",
        taskType: "pickup",
        supplierName: "Northside Suppliers",
      },
    ],
  });

  assert.equal(records.length, 1);
  assert.equal(records[0].task.type, "order");
  assert.equal(records[0].task.eventType, "order");
  assert.equal(records[0].task.title, "PICKUP: Northside Suppliers - Richmond VIC 3121");
}

{
  const records = runExport({
    tasks: [
      {
        id: "task-2",
        type: "deliver",
        orderId: "order-2",
        customerId: "cust-1",
        dueDate: "2025-01-01",
        status: "todo",
        assignedRepId: "rep-1",
      },
    ],
    orders: [
      {
        id: "order-2",
        customerId: "cust-1",
        assignedRepId: "rep-1",
        deliveryDueDate: "2025-01-01",
        status: "received",
      },
      {
        id: "order-3",
        customerId: "cust-1",
        assignedRepId: "rep-1",
        deliveryDueDate: "2025-01-01",
        status: "received",
        taskType: "pickup",
      },
    ],
  });

  assert.equal(records.length, 3);
  assert.ok(records.some((record) => record.task.eventType === "delivery"));
  assert.equal(
    records.filter((record) => record.task.eventType === "order").length,
    2
  );
}

{
  const records = runExport({
    tasks: [
      {
        id: "task-3",
        type: "deliver",
        orderId: "order-4",
        customerId: "cust-1",
        dueDate: "2025-01-01",
        status: "todo",
        assignedRepId: "rep-1",
      },
    ],
    orders: [
      {
        id: "order-4",
        customerId: "cust-1",
        assignedRepId: "rep-1",
        deliveryDueDate: "2025-01-01",
        status: "received",
        taskType: "pickup",
      },
    ],
  });

  assert.equal(records.length, 2);
  assert.ok(records.some((record) => record.task.eventType === "delivery"));
  assert.ok(records.some((record) => record.task.eventType === "order"));
}

console.log("Spokenit export tests passed.");
