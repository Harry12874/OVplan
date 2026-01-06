import {
  initDB,
  getAll,
  put,
  deleteItem,
  clearStore,
  bulkPut,
} from "./db.js";
import {
  uuid,
  todayKey,
  addDays,
  formatDate,
  formatDateTime,
  dateKeyToDate,
  parseSuburb,
  matchesSearch,
  channelLabel,
  toISO,
  dateRange,
  dayIndexFromDateKey,
  normalizeHeader,
  parseCsv,
  detectNumericColumns,
} from "./utils.js";
import {
  defaultColumns,
  fieldOptions,
  buildCsv,
  downloadCsv,
  buildSpokenitExportRecords,
  buildTaskTitle,
} from "./export.js";
import {
  supabase,
  getSession,
  signIn,
  signOut,
  supabaseAvailable,
  supabaseInitError,
} from "./supabase_client.js";

const state = {
  reps: [],
  customers: [],
  orders: [],
  tasks: [],
  scheduleEvents: [],
  oneOffItems: [],
  settings: {},
  session: null,
  selectedScheduleItems: new Set(),
  visibleScheduleItems: [],
  connection: {
    online: navigator.onLine,
    canWrite: false,
    email: "",
    status: "offline",
    cloudStatus: "offline",
    lastSyncAt: null,
  },
};

const elements = {
  appShell: document.getElementById("appShell"),
  loginScreen: document.getElementById("loginScreen"),
  loginForm: document.getElementById("loginForm"),
  loginEmail: document.getElementById("loginEmail"),
  loginPassword: document.getElementById("loginPassword"),
  loginShowPassword: document.getElementById("loginShowPassword"),
  loginError: document.getElementById("loginError"),
  loginLoading: document.getElementById("loginLoading"),
  logoutBtn: document.getElementById("logoutBtn"),
  statusIndicator: document.getElementById("statusIndicator"),
  offlineBanner: document.getElementById("offlineBanner"),
  syncErrorBanner: document.getElementById("syncErrorBanner"),
  errorScreen: document.getElementById("errorScreen"),
  errorReloadBtn: document.getElementById("errorReloadBtn"),
  errorCopyBtn: document.getElementById("errorCopyBtn"),
  errorDetails: document.getElementById("errorDetails"),
  cloudStatus: document.getElementById("cloudStatus"),
  tabs: document.querySelectorAll(".tab[data-tab]"),
  panels: document.querySelectorAll(".tab-panel"),
  dashboardRepFilter: document.getElementById("dashboardRepFilter"),
  dashboardSearch: document.getElementById("dashboardSearch"),
  todayOrdersPutList: document.getElementById("todayOrdersPutList"),
  todayOrdersGetList: document.getElementById("todayOrdersGetList"),
  todayPacksList: document.getElementById("todayPacksList"),
  todayDeliveriesList: document.getElementById("todayDeliveriesList"),
  todayOrdersPutCount: document.getElementById("todayOrdersPutCount"),
  todayOrdersGetCount: document.getElementById("todayOrdersGetCount"),
  todayPacksCount: document.getElementById("todayPacksCount"),
  todayDeliveriesCount: document.getElementById("todayDeliveriesCount"),
  addOneOffTodayBtn: document.getElementById("addOneOffTodayBtn"),
  todayToggleExpected: document.getElementById("todayToggleExpected"),
  todayTogglePacks: document.getElementById("todayTogglePacks"),
  todayToggleDeliveries: document.getElementById("todayToggleDeliveries"),
  todayStatusFilters: document.getElementById("todayStatusFilters"),
  ordersList: document.getElementById("ordersList"),
  customersList: document.getElementById("customersList"),
  repsList: document.getElementById("repsList"),
  expectedList: document.getElementById("expectedList"),
  customersSearch: document.getElementById("customersSearch"),
  customersRepFilter: document.getElementById("customersRepFilter"),
  customersModeFilter: document.getElementById("customersModeFilter"),
  customersFrequencyFilter: document.getElementById("customersFrequencyFilter"),
  customersChannelFilter: document.getElementById("customersChannelFilter"),
  customersScheduleFilter: document.getElementById("customersScheduleFilter"),
  scheduleRepFilter: document.getElementById("scheduleRepFilter"),
  scheduleToggleExpected: document.getElementById("scheduleToggleExpected"),
  scheduleTogglePacks: document.getElementById("scheduleTogglePacks"),
  scheduleToggleDeliveries: document.getElementById("scheduleToggleDeliveries"),
  scheduleStatusFilters: document.getElementById("scheduleStatusFilters"),
  scheduleList: document.getElementById("scheduleList"),
  addOneOffScheduleBtn: document.getElementById("addOneOffScheduleBtn"),
  scheduleViewButtons: document.querySelectorAll(".schedule-view-button"),
  schedulePrevBtn: document.getElementById("schedulePrevBtn"),
  scheduleNextBtn: document.getElementById("scheduleNextBtn"),
  scheduleTodayBtn: document.getElementById("scheduleTodayBtn"),
  scheduleDatePicker: document.getElementById("scheduleDatePicker"),
  scheduleRangeLabel: document.getElementById("scheduleRangeLabel"),
  scheduleSearch: document.getElementById("scheduleSearch"),
  scheduleFiltersToggle: document.getElementById("scheduleFiltersToggle"),
  scheduleFiltersPanel: document.getElementById("scheduleFiltersPanel"),
  scheduleToolbar: document.getElementById("scheduleToolbar"),
  scheduleDaySelect: document.getElementById("scheduleDaySelect"),
  scheduleDayPrev: document.getElementById("scheduleDayPrev"),
  scheduleDayNext: document.getElementById("scheduleDayNext"),
  exportStart: document.getElementById("exportStart"),
  exportEnd: document.getElementById("exportEnd"),
  exportRep: document.getElementById("exportRep"),
  exportIncludeCompleted: document.getElementById("exportIncludeCompleted"),
  exportIncludeTomorrow: document.getElementById("exportIncludeTomorrow"),
  exportSummary: document.getElementById("exportSummary"),
  mappingPresetSelect: document.getElementById("mappingPresetSelect"),
  mappingEditor: document.getElementById("mappingEditor"),
  newPresetBtn: document.getElementById("newPresetBtn"),
  savePresetBtn: document.getElementById("savePresetBtn"),
  exportCsvBtn: document.getElementById("exportCsvBtn"),
  importFile: document.getElementById("importFile"),
  importDefaultRep: document.getElementById("importDefaultRep"),
  importDuplicateMode: document.getElementById("importDuplicateMode"),
  importAovValue: document.getElementById("importAovValue"),
  importAovColumn: document.getElementById("importAovColumn"),
  importPreviewBtn: document.getElementById("importPreviewBtn"),
  importRunBtn: document.getElementById("importRunBtn"),
  importPreview: document.getElementById("importPreview"),
  importStatus: document.getElementById("importStatus"),
  importResults: document.getElementById("importResults"),
  importReportBtn: document.getElementById("importReportBtn"),
  backupBtn: document.getElementById("backupBtn"),
  restoreInput: document.getElementById("restoreInput"),
  restoreBtn: document.getElementById("restoreBtn"),
  backupStatus: document.getElementById("backupStatus"),
  newOrderBtn: document.getElementById("newOrderBtn"),
  newCustomerBtn: document.getElementById("newCustomerBtn"),
  newRepBtn: document.getElementById("newRepBtn"),
  sampleDataBtn: document.getElementById("sampleDataBtn"),
  helpBtn: document.getElementById("helpBtn"),
  modal: document.getElementById("modal"),
  modalBody: document.getElementById("modalBody"),
  modalClose: document.getElementById("modalClose"),
  snackbar: document.getElementById("snackbar"),
  snackbarMessage: document.getElementById("snackbarMessage"),
  snackbarUndo: document.getElementById("snackbarUndo"),
  accountEmail: document.getElementById("accountEmail"),
  accountLogoutBtn: document.getElementById("accountLogoutBtn"),
  scheduleSelectModeBtn: document.getElementById("scheduleSelectModeBtn"),
  scheduleSelectAllBtn: document.getElementById("scheduleSelectAllBtn"),
  scheduleExportSelectedBtn: document.getElementById("scheduleExportSelectedBtn"),
};

const statusLabels = {
  received: "Received",
  confirmed: "Confirmed",
  packed: "Packed",
  delivered: "Delivered",
  cancelled: "Cancelled",
};

function on(sel, evt, handler) {
  const el = typeof sel === "string" ? document.querySelector(sel) : sel;
  if (!el) return false;
  el.addEventListener(evt, handler);
  return true;
}

function onDoc(sel, evt, handler) {
  document.addEventListener(evt, (event) => {
    const target = event.target.closest(sel);
    if (!target) return;
    handler(event, target);
  });
}

const cadenceLabels = {
  weekly: "Weekly",
  fortnightly: "Fortnightly",
  twice_weekly: "Twice weekly",
  custom: "Custom",
  unset: "Schedule not set",
};

const dayNames = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];

const orderTermsOptions = [
  { label: "Order through portal", value: "portal" },
  { label: "SMS order", value: "sms" },
  { label: "Email order", value: "email" },
  { label: "Phone call order", value: "phone" },
  { label: "Rep in-person order", value: "rep_in_person" },
];

const orderTermsMap = orderTermsOptions.reduce((acc, option) => {
  acc[option.label.toLowerCase()] = option.value;
  return acc;
}, {});

let latestErrorDetails = "";

function formatErrorDetails(error, context) {
  const message = error?.message || String(error);
  const stack = error?.stack ? `\n\n${error.stack}` : "";
  const timestamp = new Date().toISOString();
  return `[${timestamp}] ${context}\n${message}${stack}`;
}

function showErrorScreen(error, context = "Unexpected error") {
  const details = formatErrorDetails(error, context);
  latestErrorDetails = details;
  if (elements.errorDetails) {
    elements.errorDetails.textContent = details;
  }
  if (elements.errorScreen) {
    elements.errorScreen.classList.remove("hidden");
  }
  console.error(context, error);
}

function setupGlobalErrorHandling() {
  window.addEventListener("error", (event) => {
    const error = event.error || new Error(event.message || "Unknown error");
    showErrorScreen(error, "Runtime error");
  });
  window.addEventListener("unhandledrejection", (event) => {
    const error = event.reason instanceof Error ? event.reason : new Error(String(event.reason));
    showErrorScreen(error, "Unhandled promise rejection");
  });
  if (elements.errorReloadBtn) {
    elements.errorReloadBtn.addEventListener("click", () => window.location.reload());
  }
  if (elements.errorCopyBtn) {
    elements.errorCopyBtn.addEventListener("click", async () => {
      if (!latestErrorDetails) return;
      try {
        await navigator.clipboard.writeText(latestErrorDetails);
        elements.errorCopyBtn.textContent = "Copied!";
        setTimeout(() => {
          elements.errorCopyBtn.textContent = "Copy error details";
        }, 2000);
      } catch (error) {
        console.error("Failed to copy error details", error);
      }
    });
  }
}

function orderTermsLabelFromChannel(channel) {
  return orderTermsOptions.find((option) => option.value === channel)?.label || "";
}

function scheduleFrequencyLabel(frequency) {
  if (!frequency) return cadenceLabels.unset;
  if (frequency === "WEEKLY") return "Weekly";
  if (frequency === "FORTNIGHTLY") return "Fortnightly";
  if (frequency === "EVERY_3_WEEKS") return "Every 3 weeks";
  return frequency;
}

function defaultSchedule() {
  return {
    mode: null,
    frequency: null,
    orderDay1: null,
    deliverDay1: null,
    packDay1: null,
    isBiWeeklySecondRun: false,
    orderDay2: null,
    deliverDay2: null,
    packDay2: null,
    customerOrderDays: [],
    deliverDays: [],
    packDays: [],
    anchorDate: null,
  };
}

function helpIcon(text) {
  return `<span class="help-icon" title="${text}">?</span>`;
}

function buildDayButtons(selectedDays = [], role) {
  return `
    <div class="day-buttons" data-day-role="${role}">
      ${dayNames
        .map((label, index) => {
          const selected = selectedDays.includes(index) ? "selected" : "";
          return `<button type="button" class="day-button ${selected}" data-day="${index}">${label}</button>`;
        })
        .join("")}
    </div>
  `;
}

const importState = {
  headers: [],
  rows: [],
  mapped: [],
  reportCsv: "",
  headerMap: null,
};

let lastUndo = null;
let snackbarTimeout = null;

const defaultSettings = {
  exportPresets: [
    {
      id: "default",
      name: "Default",
      columns: defaultColumns,
    },
  ],
  activeExportPresetId: "default",
  lastBackupAt: null,
  scheduleView: {
    toggles: {
      expectedOrders: true,
      packs: true,
      deliveries: true,
    },
    statusFilter: "all",
    viewMode: "week",
    anchorDate: todayKey(),
    activeDayKey: todayKey(),
    searchTerm: "",
  },
};

function ensureSettingsDefaults(settings = {}) {
  const safe = settings || {};
  return {
    ...safe,
    exportPresets: Array.isArray(safe.exportPresets) ? safe.exportPresets : [],
    exportColumnMapping: safe.exportColumnMapping || {},
    agendaToggles: safe.agendaToggles || {},
    scheduleView: {
      ...defaultSettings.scheduleView,
      ...safe.scheduleView,
      toggles: {
        ...defaultSettings.scheduleView.toggles,
        ...(safe.scheduleView?.toggles || {}),
      },
    },
  };
}

function createCustomerId() {
  if (typeof crypto !== "undefined" && crypto.randomUUID) {
    return crypto.randomUUID();
  }
  return uuid();
}

function setSyncError(message) {
  if (!elements.syncErrorBanner) return;
  elements.syncErrorBanner.textContent = message || "";
  elements.syncErrorBanner.classList.toggle("hidden", !message);
}

function updateCloudStatusText() {
  if (!elements.cloudStatus) return;
  let label = "Cloud: Offline (local only)";
  if (state.connection.cloudStatus === "syncing") {
    label = "Cloud: Syncing...";
  }
  if (state.connection.cloudStatus === "synced") {
    const timestamp = state.connection.lastSyncAt
      ? ` at ${formatDateTime(state.connection.lastSyncAt)}`
      : "";
    label = `Cloud: Synced${timestamp}`;
  }
  elements.cloudStatus.textContent = label;
}

function setCloudStatus(status, timestamp = null) {
  state.connection.cloudStatus = status;
  if (timestamp) {
    state.connection.lastSyncAt = timestamp;
  }
  updateCloudStatusText();
}

function isNetworkFailure(error) {
  if (!navigator.onLine) return true;
  const message = error?.message || "";
  return (
    error?.status === 0 ||
    message.includes("Failed to fetch") ||
    message.includes("NetworkError") ||
    message.includes("Network request failed")
  );
}

function handleSupabaseError(error, { context = "Supabase error", alertOnOffline = false } = {}) {
  const message = error?.message ? `${context}: ${error.message}` : context;
  if (isNetworkFailure(error)) {
    updateConnectionStatus({
      online: false,
      canWrite: false,
      status: "Offline/Local",
    });
    setSyncError("");
    setCloudStatus("offline");
    if (alertOnOffline) {
      showOfflineAlert();
    }
    return;
  }
  updateConnectionStatus({
    online: true,
    canWrite: true,
    status: "Sync error",
  });
  setSyncError(message);
  updateCloudStatusText();
}

function updateConnectionStatus({ online, canWrite, email, status }) {
  state.connection.online = online ?? state.connection.online;
  state.connection.canWrite = canWrite ?? state.connection.canWrite;
  state.connection.email = email ?? state.connection.email;
  state.connection.status = status ?? state.connection.status;
  if (status === "Online/Synced") {
    setSyncError("");
  }
  if (status === "Offline/Local" || !state.connection.canWrite) {
    setCloudStatus("offline");
  }
  if (elements.statusIndicator) {
    elements.statusIndicator.textContent = state.connection.email
      ? `Logged in as ${state.connection.email} • ${state.connection.status}`
      : "Not logged in";
  }
  if (elements.accountEmail) {
    elements.accountEmail.textContent = state.connection.email || "Not logged in";
  }
  if (elements.offlineBanner) {
    elements.offlineBanner.classList.toggle(
      "hidden",
      !state.session || state.connection.canWrite
    );
  }
  document.querySelectorAll("[data-requires-online='true']").forEach((element) => {
    element.disabled = !state.connection.canWrite;
  });
}

function canWrite() {
  return state.connection.canWrite;
}

function showLoginScreen() {
  if (elements.appShell) elements.appShell.classList.add("hidden");
  if (elements.loginScreen) elements.loginScreen.classList.remove("hidden");
  if (elements.loginForm) elements.loginForm.reset();
  setLoginError("");
  setSyncError("");
  setLoginLoading(false);
  setCloudStatus("offline");
  updateConnectionStatus({
    canWrite: false,
    status: "offline",
    email: "",
  });
}

function showAppShell() {
  if (elements.loginScreen) elements.loginScreen.classList.add("hidden");
  if (elements.appShell) elements.appShell.classList.remove("hidden");
}

function setLoginLoading(isLoading) {
  if (elements.loginLoading) {
    elements.loginLoading.classList.toggle("hidden", !isLoading);
  }
  if (elements.loginForm) {
    elements.loginForm.querySelector("button[type='submit']").disabled = isLoading;
  }
}

function setLoginError(message) {
  if (elements.loginError) {
    elements.loginError.textContent = message || "";
    elements.loginError.classList.toggle("hidden", !message);
  }
}

function showOfflineAlert() {
  alert("Offline: connect to sync changes.");
}

function disableFormIfOffline(formId) {
  const form = document.getElementById(formId);
  if (!form || canWrite()) return;
  const notice = document.createElement("div");
  notice.className = "offline-message";
  notice.textContent = "Offline: changes can’t sync yet.";
  form.prepend(notice);
  form.querySelectorAll("button").forEach((button) => {
    if (button.type === "submit" || button.dataset.action) {
      button.disabled = true;
    }
  });
}

async function loadState({ useLocalCustomers = true, useLocalScheduleEvents = true } = {}) {
  state.reps = await getAll("reps");
  state.customers = useLocalCustomers ? await getAll("customers") : [];
  state.orders = await getAll("orders");
  state.tasks = await getAll("tasks");
  state.scheduleEvents = useLocalScheduleEvents ? await getAll("schedule_events") : [];
  state.oneOffItems = await getAll("one_off_items");
  const settings = await getAll("settings");
  state.settings = ensureSettingsDefaults(
    settings.reduce((acc, item) => {
      acc[item.id] = item;
      return acc;
    }, {})
  );

  if (!state.settings.app) {
    state.settings.app = ensureSettingsDefaults({ id: "app", ...defaultSettings });
    await put("settings", state.settings.app);
  }
  if (!state.settings.app.scheduleView) {
    state.settings.app.scheduleView = { ...defaultSettings.scheduleView };
    await saveSettings();
  } else {
    state.settings.app.scheduleView.toggles = {
      ...defaultSettings.scheduleView.toggles,
      ...state.settings.app.scheduleView.toggles,
    };
    state.settings.app.scheduleView.statusFilter =
      state.settings.app.scheduleView.statusFilter || defaultSettings.scheduleView.statusFilter;
    state.settings.app.scheduleView.viewMode =
      state.settings.app.scheduleView.viewMode || defaultSettings.scheduleView.viewMode;
    state.settings.app.scheduleView.anchorDate =
      state.settings.app.scheduleView.anchorDate || defaultSettings.scheduleView.anchorDate;
    state.settings.app.scheduleView.searchTerm =
      state.settings.app.scheduleView.searchTerm ?? defaultSettings.scheduleView.searchTerm;
  }

  let customersUpdated = false;
  state.customers = state.customers.map((customer) => {
    const updated = { ...customer };
    let touched = false;
    if (!updated.customerId) {
      updated.customerId = updated.externalCustomerId || "";
      touched = true;
    }
    if (!updated.orderChannel && updated.channelPreference) {
      updated.orderChannel = updated.channelPreference;
      touched = true;
    }
    if (!updated.orderTermsLabel && updated.orderChannel) {
      updated.orderTermsLabel = orderTermsLabelFromChannel(updated.orderChannel);
      touched = true;
    }
    if (!updated.extraFields) {
      updated.extraFields = {};
      touched = true;
    }
    if (!updated.schedule) {
      updated.schedule = defaultSchedule();
      touched = true;
    }
    if (!updated.customerNotes) {
      updated.customerNotes = "";
      touched = true;
    }
    if (updated.schedule) {
      if (updated.schedule.packDay1 === undefined) {
        updated.schedule.packDay1 = null;
        touched = true;
      }
      if (updated.schedule.packDay2 === undefined) {
        updated.schedule.packDay2 = null;
        touched = true;
      }
      if (!updated.schedule.packDays) {
        updated.schedule.packDays = [];
        touched = true;
      }
    }
    if (updated.cadenceType === undefined) {
      updated.cadenceType = null;
      touched = true;
    }
    if (touched) {
      customersUpdated = true;
      return updated;
    }
    return customer;
  });

  if (customersUpdated) {
    for (const customer of state.customers) {
      await put("customers", customer);
    }
  }
}

function saveSettings() {
  return put("settings", state.settings.app);
}

function showModal(contentHtml) {
  elements.modalBody.innerHTML = contentHtml;
  elements.modal.classList.remove("hidden");
}

function closeModal() {
  elements.modal.classList.add("hidden");
  elements.modalBody.innerHTML = "";
}

function showSnackbar(message, undoFn) {
  elements.snackbarMessage.textContent = message;
  elements.snackbar.classList.remove("hidden");
  lastUndo = undoFn;
  if (snackbarTimeout) {
    clearTimeout(snackbarTimeout);
  }
  snackbarTimeout = setTimeout(() => {
    elements.snackbar.classList.add("hidden");
    lastUndo = null;
  }, 6000);
}

function hideSnackbar() {
  elements.snackbar.classList.add("hidden");
  lastUndo = null;
}

function getRepOptions(includeAll = true) {
  const options = includeAll ? [{ id: "all", name: "All reps" }] : [];
  return options.concat(state.reps.filter((rep) => rep.active !== false));
}

function repName(repId) {
  return state.reps.find((rep) => rep.id === repId)?.name || "Unassigned";
}

function customerById(id) {
  return state.customers.find((customer) => customer.id === id);
}

function orderById(id) {
  return state.orders.find((order) => order.id === id);
}

function tasksForOrder(orderId) {
  return state.tasks.filter((task) => task.orderId === orderId);
}

function computeDueDates(receivedAt, packOffsetDays, deliveryOffsetDays) {
  const receivedDateKey = todayKeyFromISO(receivedAt);
  return {
    packDueDate: addDays(receivedDateKey, packOffsetDays),
    deliveryDueDate: addDays(receivedDateKey, deliveryOffsetDays),
  };
}

function todayKeyFromISO(isoString) {
  return todayKeyFromDate(new Date(isoString));
}

function todayKeyFromDate(date) {
  const parts = new Intl.DateTimeFormat("en-AU", {
    timeZone: "Australia/Melbourne",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).formatToParts(date);
  const lookup = Object.fromEntries(parts.map((p) => [p.type, p.value]));
  return `${lookup.year}-${lookup.month}-${lookup.day}`;
}

function buildTask(order, type, dueDate) {
  return {
    id: uuid(),
    type,
    orderId: order.id,
    customerId: order.customerId,
    dueDate,
    status: "todo",
    assignedRepId: order.assignedRepId,
  };
}

async function createOrderTasks(order) {
  const packTask = buildTask(order, "pack", order.packDueDate);
  const deliveryTask = buildTask(order, "deliver", order.deliveryDueDate);
  state.tasks.push(packTask, deliveryTask);
  await bulkPut("tasks", [packTask, deliveryTask]);
}

async function removeTasksForOrder(orderId) {
  const tasksToDelete = state.tasks.filter((task) => task.orderId === orderId);
  for (const task of tasksToDelete) {
    await deleteItem("tasks", task.id);
  }
  state.tasks = state.tasks.filter((task) => task.orderId !== orderId);
}

function renderTabs() {
  elements.tabs.forEach((tab) => {
    tab.addEventListener("click", () => {
      elements.tabs.forEach((t) => t.classList.remove("active"));
      const target = tab.dataset.tab;
      elements.tabs.forEach((t) => {
        if (t.dataset.tab === target) {
          t.classList.add("active");
        }
      });
      elements.panels.forEach((panel) => {
        panel.classList.toggle("active", panel.id === target);
      });
    });
  });
}

function renderRepFilters() {
  const repOptions = getRepOptions();
  elements.dashboardRepFilter.innerHTML = repOptions
    .map((rep) => `<option value="${rep.id}">${rep.name}</option>`)
    .join("");

  elements.exportRep.innerHTML = repOptions
    .map((rep) => `<option value="${rep.id}">${rep.name}</option>`)
    .join("");

  elements.importDefaultRep.innerHTML = [
    "<option value=\"\">Select rep</option>",
    ...state.reps.map((rep) => `<option value="${rep.id}">${rep.name}</option>`),
  ].join("");

  elements.customersRepFilter.innerHTML = repOptions
    .map((rep) => `<option value="${rep.id}">${rep.name}</option>`)
    .join("");

  elements.scheduleRepFilter.innerHTML = repOptions
    .map((rep) => `<option value="${rep.id}">${rep.name}</option>`)
    .join("");
}

function setStatusChipSelection(container, value) {
  if (!container) return;
  container.querySelectorAll(".filter-chip").forEach((chip) => {
    chip.classList.toggle("active", chip.dataset.status === value);
  });
}

function syncScheduleViewControls() {
  const view = state.settings.app.scheduleView;
  if (elements.scheduleToggleExpected) {
    elements.scheduleToggleExpected.checked = view.toggles.expectedOrders;
    elements.scheduleTogglePacks.checked = view.toggles.packs;
    elements.scheduleToggleDeliveries.checked = view.toggles.deliveries;
  }
  if (elements.todayToggleExpected) {
    elements.todayToggleExpected.checked = view.toggles.expectedOrders;
    elements.todayTogglePacks.checked = view.toggles.packs;
    elements.todayToggleDeliveries.checked = view.toggles.deliveries;
  }
  setStatusChipSelection(elements.scheduleStatusFilters, view.statusFilter);
  setStatusChipSelection(elements.todayStatusFilters, view.statusFilter);
  if (elements.scheduleSearch) {
    elements.scheduleSearch.value = view.searchTerm || "";
  }
  if (elements.scheduleViewButtons?.length) {
    elements.scheduleViewButtons.forEach((button) => {
      button.classList.toggle("active", button.dataset.view === view.viewMode);
    });
  }
  if (elements.scheduleDatePicker) {
    elements.scheduleDatePicker.value = view.anchorDate || todayKey();
  }
}

function scheduleEventLookupKey(event) {
  if (event.sourceId) {
    return `oneoff-${event.sourceId}-${event.date}-${event.kind}`;
  }
  return `${event.customerId}-${event.kind}-${event.date}-${event.runIndex ?? 0}`;
}

function scheduleEventKeyForItem(item) {
  if (item.sourceId) {
    return `oneoff-${item.sourceId}-${item.date}-${item.kind}`;
  }
  return `${item.customerId}-${item.kind}-${item.date}-${item.runIndex ?? 0}`;
}

function agendaKindLabel(item) {
  if (item.kind === "custom_oneoff") {
    if (item.oneOffKind === "delivery") return "One-off delivery";
    if (item.oneOffKind === "pack") return "One-off pack";
    return "One-off expected order";
  }
  if (item.kind === "delivery") return "Delivery";
  if (item.kind === "pack") return "Pack";
  return "Expected order";
}

function agendaTypeText(item) {
  const baseKind = item.kind === "custom_oneoff" ? item.oneOffKind : item.kind;
  if (baseKind === "delivery") return "DELIVERY";
  if (baseKind === "pack") return "PACKING";
  return "EXPECTED ORDER";
}

function agendaTypeClass(item) {
  const baseKind = item.kind === "custom_oneoff" ? item.oneOffKind : item.kind;
  if (baseKind === "delivery") return "type-delivery";
  if (baseKind === "pack") return "type-pack";
  return "type-order";
}

function agendaTypeLabelClass(item) {
  const baseKind = item.kind === "custom_oneoff" ? item.oneOffKind : item.kind;
  if (baseKind === "delivery") return "delivery";
  if (baseKind === "pack") return "pack";
  return "order";
}

function agendaActionLabel(item) {
  const kind = item.kind === "custom_oneoff" ? item.oneOffKind : item.kind;
  if (kind === "delivery") return "Mark Delivered";
  if (kind === "pack") return "Mark Packed";
  return "Mark Order Taken";
}

function agendaSortOrder(item) {
  const kind = item.kind === "custom_oneoff" ? item.oneOffKind : item.kind;
  if (kind === "delivery") return 1;
  if (kind === "pack") return 2;
  if (kind === "expected_order") return 3;
  return 4;
}

function nextDateForWeekday(startKey, targetDayIndex) {
  const current = dayIndexFromDateKey(startKey);
  let diff = targetDayIndex - current;
  if (diff < 0) diff += 7;
  return addDays(startKey, diff);
}

function nextDateForAnyWeekday(startKey, dayIndices = []) {
  const dates = dayIndices.map((dayIndex) => nextDateForWeekday(startKey, dayIndex));
  return dates.sort()[0] || startKey;
}

function scheduleAppliesOnDate(schedule, dateKey) {
  if (!schedule?.frequency) return false;
  if (schedule.frequency === "WEEKLY") return true;
  if (schedule.frequency === "FORTNIGHTLY") {
    return isWeekAligned(dateKey, schedule.anchorDate, 2);
  }
  if (schedule.frequency === "EVERY_3_WEEKS") {
    return isWeekAligned(dateKey, schedule.anchorDate, 3);
  }
  return false;
}

function isWeekAligned(dateKey, anchorDate, cycleWeeks) {
  if (!anchorDate) return false;
  const date = dateKeyToDate(dateKey);
  const anchor = dateKeyToDate(anchorDate);
  const diffDays = Math.floor((date - anchor) / (1000 * 60 * 60 * 24));
  const weeks = Math.floor(diffDays / 7);
  const normalized = ((weeks % cycleWeeks) + cycleWeeks) % cycleWeeks;
  return normalized === 0;
}

function rangeKeys(startKey, endKey) {
  const start = dateKeyToDate(startKey);
  const end = dateKeyToDate(endKey);
  const days = Math.floor((end - start) / (1000 * 60 * 60 * 24)) + 1;
  return dateRange(startKey, Math.max(days, 1));
}

function buildScheduleItems(startKey, endKey) {
  const items = [];
  const seen = new Set();
  const keys = rangeKeys(startKey, endKey);

  const addItem = (item) => {
    const key = `${item.customerId}-${item.kind}-${item.date}-${item.runIndex ?? "x"}`;
    if (seen.has(key)) return;
    seen.add(key);
    items.push(item);
  };

  keys.forEach((dateKey) => {
    const dayIndex = dayIndexFromDateKey(dateKey);
    state.customers.forEach((customer) => {
      const schedule = customer.schedule;
      if (!schedule?.mode || !schedule.frequency) return;
      if (!scheduleAppliesOnDate(schedule, dateKey)) return;

      if (schedule.mode === "WE_GET_ORDER") {
        const runs = [
          {
            orderDay: schedule.orderDay1,
            deliverDay: schedule.deliverDay1,
            packDay: schedule.packDay1,
            runIndex: 1,
          },
        ];
        if (schedule.isBiWeeklySecondRun) {
          runs.push({
            orderDay: schedule.orderDay2,
            deliverDay: schedule.deliverDay2,
            packDay: schedule.packDay2,
            runIndex: 2,
          });
        }
        runs.forEach((run) => {
          if (run.orderDay !== null && run.orderDay !== undefined && dayIndex === run.orderDay) {
            addItem({
              id: uuid(),
              kind: "expected_order",
              date: dateKey,
              customerId: customer.id,
              repId: customer.assignedRepId,
              runIndex: run.runIndex,
              frequency: schedule.frequency,
              title: customer.storeName,
              subtitle: customer.contactName || "",
              orderMode: schedule.mode,
            });
          }

          const packDay = run.packDay ?? run.orderDay;
          if (packDay !== null && packDay !== undefined && dayIndex === packDay) {
            addItem({
              id: uuid(),
              kind: "pack",
              date: dateKey,
              customerId: customer.id,
              repId: customer.assignedRepId,
              runIndex: run.runIndex,
              frequency: schedule.frequency,
              title: customer.storeName,
              subtitle: customer.contactName || "",
              orderMode: schedule.mode,
            });
          }

          const deliverDay =
            run.deliverDay ?? (run.orderDay !== null && run.orderDay !== undefined ? (run.orderDay + 1) % 7 : null);
          if (deliverDay !== null && deliverDay !== undefined && dayIndex === deliverDay) {
            addItem({
              id: uuid(),
              kind: "delivery",
              date: dateKey,
              customerId: customer.id,
              repId: customer.assignedRepId,
              runIndex: run.runIndex,
              frequency: schedule.frequency,
              title: customer.storeName,
              subtitle: customer.contactName || "",
              orderMode: schedule.mode,
            });
          }
        });
      }

      if (schedule.mode === "THEY_PUT_ORDER") {
        schedule.customerOrderDays.forEach((orderDay) => {
          if (dayIndex !== orderDay) return;
          addItem({
            id: uuid(),
            kind: "expected_order",
            date: dateKey,
            customerId: customer.id,
            repId: customer.assignedRepId,
            runIndex: null,
            frequency: schedule.frequency,
            title: customer.storeName,
            subtitle: customer.contactName || "",
            orderMode: schedule.mode,
          });
        });

        const packDays =
          schedule.packDays && schedule.packDays.length ? schedule.packDays : schedule.customerOrderDays;
        if (packDays?.includes(dayIndex)) {
          addItem({
            id: uuid(),
            kind: "pack",
            date: dateKey,
            customerId: customer.id,
            repId: customer.assignedRepId,
            runIndex: null,
            frequency: schedule.frequency,
            title: customer.storeName,
            subtitle: customer.contactName || "",
            orderMode: schedule.mode,
          });
        }

        const deliverDays =
          schedule.deliverDays && schedule.deliverDays.length
            ? schedule.deliverDays
            : schedule.customerOrderDays.map((orderDay) => (orderDay + 1) % 7);
        if (deliverDays?.includes(dayIndex)) {
          addItem({
            id: uuid(),
            kind: "delivery",
            date: dateKey,
            customerId: customer.id,
            repId: customer.assignedRepId,
            runIndex: null,
            frequency: schedule.frequency,
            title: customer.storeName,
            subtitle: customer.contactName || "",
            orderMode: schedule.mode,
          });
        }
      }
    });
  });

  return items;
}

function buildAgendaItems({ dateStart, dateEnd, toggles, statusFilter, repFilter, searchTerm }) {
  const items = [];
  items.push(...buildScheduleItems(dateStart, dateEnd));

  state.oneOffItems
    .filter((item) => !item.isDeleted)
    .filter((item) => item.date >= dateStart && item.date <= dateEnd)
    .forEach((item) => {
      const customer = item.customerId ? customerById(item.customerId) : null;
      items.push({
        id: uuid(),
        kind: "custom_oneoff",
        oneOffKind: item.kind,
        date: item.date,
        customerId: item.customerId,
        repId: item.repId || customer?.assignedRepId || "",
        runIndex: null,
        sourceId: item.id,
        title: customer?.storeName || item.note || "One-off item",
        subtitle: customer?.contactName || "",
        orderMode: customer?.schedule?.mode || null,
        note: item.note,
        isOneOff: true,
      });
    });

  const eventMap = new Map(
    state.scheduleEvents.map((event) => [scheduleEventLookupKey(event), event])
  );

  const filtered = items
    .map((item) => {
      const event = eventMap.get(scheduleEventKeyForItem(item));
      const status = event?.status || "PENDING";
      const selectionId = scheduleEventKeyForItem(item);
      return {
        ...item,
        status,
        selectionId,
        skippedReason: event?.skippedReason || null,
        skippedReasonText: event?.skippedReasonText || null,
        completedAt: event?.completedAt || null,
      };
    })
    .filter((item) => {
      const baseKind = item.kind === "custom_oneoff" ? item.oneOffKind : item.kind;
      if (baseKind === "expected_order" && !toggles.expectedOrders) return false;
      if (baseKind === "pack" && !toggles.packs) return false;
      if (baseKind === "delivery" && !toggles.deliveries) return false;
      if (statusFilter && statusFilter !== "all" && item.status.toLowerCase() !== statusFilter) {
        return false;
      }
      if (repFilter && repFilter !== "all" && item.repId !== repFilter) return false;
      if (searchTerm) {
        const customer = item.customerId ? customerById(item.customerId) : null;
        const combined = [
          item.title,
          item.subtitle,
          customer?.fullAddress,
          customer?.phone,
          customer?.email,
        ]
          .filter(Boolean)
          .join(" ");
        return matchesSearch(combined, searchTerm);
      }
      return true;
    })
    .sort((a, b) => {
      if (a.date !== b.date) return a.date.localeCompare(b.date);
      const diff = agendaSortOrder(a) - agendaSortOrder(b);
      if (diff !== 0) return diff;
      return a.title.localeCompare(b.title);
    });

  return filtered;
}

function getItemsForRange({ dateStart, dateEnd, repFilter, statusFilter, toggles, searchTerm }) {
  return buildAgendaItems({
    dateStart,
    dateEnd,
    toggles,
    statusFilter,
    repFilter,
    searchTerm,
  });
}

function attachScheduleSelectionHandlers(container, items) {
  const checkboxNodes = container.querySelectorAll("input[data-selection-id]");
  checkboxNodes.forEach((checkbox) => {
    checkbox.addEventListener("click", (event) => {
      event.stopPropagation();
    });
    checkbox.addEventListener("change", (event) => {
      const id = event.target.dataset.selectionId;
      if (!id) return;
      if (event.target.checked) {
        state.selectedScheduleItems.add(id);
      } else {
        state.selectedScheduleItems.delete(id);
      }
    });
  });
  state.visibleScheduleItems = items;
}

function selectAllVisibleScheduleItems() {
  state.visibleScheduleItems.forEach((item) => state.selectedScheduleItems.add(item.selectionId));
  elements.scheduleList
    .querySelectorAll("input[data-selection-id]")
    .forEach((checkbox) => {
      checkbox.checked = true;
    });
}

function clearScheduleSelection() {
  state.selectedScheduleItems.clear();
  elements.scheduleList
    .querySelectorAll("input[data-selection-id]")
    .forEach((checkbox) => {
      checkbox.checked = false;
    });
}

function exportSelectedScheduleItems() {
  const selectedItems = state.visibleScheduleItems.filter((item) =>
    state.selectedScheduleItems.has(item.selectionId)
  );
  const deliveryItems = selectedItems.filter(
    (item) => (item.kind === "custom_oneoff" ? item.oneOffKind : item.kind) === "delivery"
  );
  if (!deliveryItems.length) {
    alert("Select at least one delivery to export.");
    return;
  }
  const preset = getActivePreset();
  if (!preset) return;
  const records = deliveryItems.map((item) => {
    const customer = customerById(item.customerId) || {};
    const rep = state.reps.find((entry) => entry.id === item.repId) || {};
    return {
      task: {
        dueDate: item.date,
        assignedRepId: item.repId || "",
        title: buildTaskTitle({ kind: "delivery", customer, order: {} }),
      },
      order: {},
      customer,
      rep,
    };
  });
  const csv = buildCsv(records, preset.columns);
  downloadCsv(`spoke-export-${todayKey()}.csv`, csv);
}

function startOfWeekMonday(dateKey) {
  const date = dateKeyToDate(dateKey);
  const day = date.getDay();
  const diff = day === 0 ? -6 : 1 - day;
  const start = new Date(date);
  start.setDate(start.getDate() + diff);
  return todayKeyFromDate(start);
}

function startOfMonth(dateKey) {
  const date = dateKeyToDate(dateKey);
  const first = new Date(date.getFullYear(), date.getMonth(), 1);
  return todayKeyFromDate(first);
}

function endOfMonth(dateKey) {
  const date = dateKeyToDate(dateKey);
  const last = new Date(date.getFullYear(), date.getMonth() + 1, 0);
  return todayKeyFromDate(last);
}

function monthGridRange(dateKey) {
  const monthStart = startOfMonth(dateKey);
  const monthEnd = endOfMonth(dateKey);
  const gridStart = startOfWeekMonday(monthStart);
  const endDay = dayIndexFromDateKey(monthEnd);
  const trailing = endDay === 0 ? 0 : 7 - endDay;
  const gridEnd = addDays(monthEnd, trailing);
  return { gridStart, gridEnd };
}

function addMonthsToDateKey(dateKey, deltaMonths) {
  const date = dateKeyToDate(dateKey);
  const year = date.getFullYear();
  const month = date.getMonth() + deltaMonths;
  const day = date.getDate();
  const candidate = new Date(year, month, 1);
  const lastDay = new Date(candidate.getFullYear(), candidate.getMonth() + 1, 0).getDate();
  const finalDate = new Date(candidate.getFullYear(), candidate.getMonth(), Math.min(day, lastDay));
  return todayKeyFromDate(finalDate);
}

function formatRangeLabel(viewMode, anchorDate) {
  const date = dateKeyToDate(anchorDate);
  const dayFormatter = new Intl.DateTimeFormat("en-AU", {
    timeZone: "Australia/Melbourne",
    day: "numeric",
    month: "short",
    year: "numeric",
  });
  const monthFormatter = new Intl.DateTimeFormat("en-AU", {
    timeZone: "Australia/Melbourne",
    month: "short",
    year: "numeric",
  });
  if (viewMode === "month") {
    return monthFormatter.format(date);
  }
  if (viewMode === "day") {
    return dayFormatter.format(date);
  }
  const start = startOfWeekMonday(anchorDate);
  const end = addDays(start, 6);
  const startDate = dateKeyToDate(start);
  const endDate = dateKeyToDate(end);
  const rangeFormatter = new Intl.DateTimeFormat("en-AU", {
    timeZone: "Australia/Melbourne",
    day: "numeric",
    month: "short",
  });
  const startLabel = rangeFormatter.format(startDate);
  const endLabel = new Intl.DateTimeFormat("en-AU", {
    timeZone: "Australia/Melbourne",
    day: "numeric",
    month: "short",
    year: "numeric",
  }).format(endDate);
  return `${startLabel}–${endLabel}`;
}

function mapCustomerToSupabase(customer) {
  return {
    id: customer.id,
    data_json: customer,
    updated_at: toISO(),
  };
}

function mapCustomerFromSupabase(row) {
  const data = row?.data_json || {};
  return {
    ...data,
    id: row?.id,
    schedule: data.schedule || defaultSchedule(),
  };
}

function mapScheduleEventToSupabase(event) {
  return {
    id: event.id,
    data_json: event,
    updated_at: toISO(),
  };
}

function mapScheduleEventFromSupabase(row) {
  const data = row?.data_json || {};
  return {
    ...data,
    id: row?.id,
  };
}

async function syncUpsertBatch(table, records, mapper, { onProgress } = {}) {
  const batchSize = 200;
  let processed = 0;
  for (let start = 0; start < records.length; start += batchSize) {
    const chunk = records.slice(start, start + batchSize).map(mapper);
    if (!chunk.length) continue;
    setCloudStatus("syncing");
    const { error } = await supabase.from(table).upsert(chunk, {
      onConflict: "id",
    });
    if (error) throw error;
    processed += chunk.length;
    if (onProgress) {
      onProgress(processed, records.length);
    }
  }
}

async function loadFromSupabase() {
  if (!state.session) return;
  setCloudStatus("syncing");
  try {
    const { data: customersData, error: customersError } = await supabase
      .from("customers")
      .select("id,data_json");
    if (customersError) throw customersError;
    const { data: eventsData, error: eventsError } = await supabase
      .from("schedule_events")
      .select("id,data_json");
    if (eventsError) throw eventsError;
    const customers = (customersData || []).map((row) => {
      const customer = mapCustomerFromSupabase(row);
      return customer.id ? customer : { ...customer, id: createCustomerId() };
    });
    const events = (eventsData || []).map(mapScheduleEventFromSupabase);
    await clearStore("customers");
    await clearStore("schedule_events");
    await bulkPut("customers", customers);
    await bulkPut("schedule_events", events);
    updateConnectionStatus({
      online: true,
      canWrite: true,
      status: "Online/Synced",
    });
    setCloudStatus("synced", toISO());
    return { status: "success" };
  } catch (error) {
    console.error("Supabase sync failed", error);
    handleSupabaseError(error, { context: "Supabase sync failed", alertOnOffline: false });
    if (isNetworkFailure(error)) {
      return { status: "network" };
    }
    return { status: "error" };
  }
}

async function syncUpsertCustomer(customer) {
  const payload = mapCustomerToSupabase(customer);
  setCloudStatus("syncing");
  const { error } = await supabase.from("customers").upsert(payload, {
    onConflict: "id",
  });
  if (!error) return;
  throw error;
}

async function syncUpsertScheduleEvent(event) {
  const payload = mapScheduleEventToSupabase(event);
  setCloudStatus("syncing");
  const { error } = await supabase.from("schedule_events").upsert(payload, {
    onConflict: "id",
  });
  if (!error) return;
  throw error;
}

async function syncDeleteScheduleEvent(eventId) {
  setCloudStatus("syncing");
  const { error } = await supabase.from("schedule_events").delete().eq("id", eventId);
  if (error) throw error;
}

async function syncPushAll() {
  if (!state.session) return;
  await syncUpsertBatch("customers", state.customers, mapCustomerToSupabase);
  await syncUpsertBatch("schedule_events", state.scheduleEvents, mapScheduleEventToSupabase);
  setCloudStatus("synced", toISO());
}

function agendaItemActions(item) {
  if (item.status === "PENDING") {
    return `
      <button class="secondary" data-action="done">${agendaActionLabel(item)}</button>
      ${item.isOneOff ? '<button class="ghost" data-action="bump-back">Bump back 1 day</button>' : ""}
      <button class="ghost" data-action="skip">Skip</button>
      ${item.isOneOff ? '<button class="ghost danger" data-action="delete">Delete</button>' : ""}
    `;
  }
  return `
    <button class="secondary" data-action="undo">Undo</button>
    ${item.isOneOff ? '<button class="ghost danger" data-action="delete">Delete</button>' : ""}
  `;
}

function agendaItemCard(item) {
  const customer = item.customerId ? customerById(item.customerId) : null;
  const suburb = customer?.suburb1 || parseSuburb(customer?.fullAddress || "");
  const postcode = customer?.postcode1 || "";
  const frequencyLabel = item.frequency ? scheduleFrequencyLabel(item.frequency) : "";
  const deliveryDetails =
    item.kind === "delivery" || item.oneOffKind === "delivery"
      ? `<div class="muted">${customer?.fullAddress || ""}</div>
         <div class="muted">${customer?.deliveryNotes || ""}</div>`
      : "";
  const sourceLabel = itemOrderSource(item);
  const repBadge =
    (item.kind === "delivery" || item.oneOffKind === "delivery") && item.repId
      ? `<span class="badge">${repName(item.repId)}</span>`
      : "";
  const titleElement = item.customerId
    ? `<button class="link-button" data-action="open">${item.title}</button>`
    : `<span>${item.title}</span>`;
  return `
    <div class="list-item agenda-item ${item.status.toLowerCase()} ${agendaTypeClass(item)}" data-item-id="${item.selectionId}">
      <header>
        <div class="agenda-header-row">
          <span class="type-label ${agendaTypeLabelClass(item)}">${agendaTypeText(item)}</span>
        </div>
        <div class="agenda-header-row">
          ${titleElement}
          <span class="agenda-meta">— ${sourceLabel}</span>
        </div>
        <div class="badges">
          ${frequencyLabel ? `<span class="badge">${frequencyLabel}</span>` : ""}
          ${repBadge}
        </div>
      </header>
      ${item.subtitle ? `<div class="muted">${item.subtitle}</div>` : ""}
      ${suburb || postcode ? `<div class="muted">${suburb} ${postcode}</div>` : ""}
      ${deliveryDetails}
      ${item.note ? `<div class="muted">${item.note}</div>` : ""}
      ${item.status === "SKIPPED" ? `<div class="status-badge">Skipped: ${item.skippedReasonText || item.skippedReason}</div>` : ""}
      ${item.status === "DONE" ? `<div class="status-badge ok">Done</div>` : ""}
      <div class="form-actions">
        ${agendaItemActions(item)}
      </div>
    </div>
  `;
}

function scheduleItemDetailCard(item, { showActions = true } = {}) {
  const customer = item.customerId ? customerById(item.customerId) : null;
  const suburb = customer?.suburb1 || parseSuburb(customer?.fullAddress || "");
  const postcode = customer?.postcode1 || "";
  const sourceLabel = itemOrderSource(item);
  const typeClass = agendaTypeLabelClass(item);
  const repLabel =
    item.kind === "pack" || item.oneOffKind === "pack" || !item.repId ? "" : repName(item.repId);
  const statusLabel =
    item.status === "DONE" ? "Done" : item.status === "SKIPPED" ? "Skipped" : "Pending";
  const orderModeLabel = scheduleOrderModeLabel(item);
  const selectionChecked = state.selectedScheduleItems.has(item.selectionId) ? "checked" : "";
  return `
    <div class="schedule-detail-card ${agendaTypeClass(item)} is-collapsed" data-item-id="${item.selectionId}">
      <header class="schedule-detail-header">
        <div class="schedule-detail-header-left">
          <label class="schedule-select">
            <input type="checkbox" data-selection-id="${item.selectionId}" ${selectionChecked} />
            <span class="sr-only">Select</span>
          </label>
          <div class="type-label ${typeClass}">${agendaTypeText(item)}</div>
        </div>
        <span class="status-pill ${item.status.toLowerCase()}">${statusLabel}</span>
      </header>
      <div class="schedule-detail-title">
        <button class="link-button" data-action="open">${item.title}</button>
        <span class="agenda-meta">— ${sourceLabel}${orderModeLabel ? ` (${orderModeLabel})` : ""}</span>
      </div>
      <div class="schedule-detail-body">
        ${item.subtitle ? `<div class="muted">${item.subtitle}</div>` : ""}
        ${suburb || postcode ? `<div class="muted">${suburb} ${postcode}</div>` : ""}
        ${repLabel ? `<div class="muted">Rep: ${repLabel}</div>` : ""}
        ${item.note ? `<div class="muted">${item.note}</div>` : ""}
        ${item.status === "SKIPPED" ? `<div class="status-badge">Skipped: ${item.skippedReasonText || item.skippedReason}</div>` : ""}
      </div>
      <div class="schedule-detail-actions">
        <button class="btn btn-secondary btn-small" type="button" data-action="toggle-details" aria-expanded="false">Expand</button>
        ${showActions ? `<div class="form-actions">${agendaItemActions(item)}</div>` : ""}
      </div>
    </div>
  `;
}

function scheduleItemChip(item) {
  const selectionChecked = state.selectedScheduleItems.has(item.selectionId) ? "checked" : "";
  const orderModeLabel = scheduleOrderModeLabel(item);
  const statusLabel =
    item.status === "DONE" ? "Done" : item.status === "SKIPPED" ? "Skipped" : "Pending";
  return `
    <div class="schedule-chip-row ${agendaTypeClass(item)}">
      <label class="schedule-select">
        <input type="checkbox" data-selection-id="${item.selectionId}" ${selectionChecked} />
        <span class="sr-only">Select</span>
      </label>
      <button class="schedule-chip" data-action="details">
        <span class="chip-dot ${agendaTypeLabelClass(item)}"></span>
        <span class="chip-title">${item.title}</span>
        <span class="status-pill ${item.status.toLowerCase()} chip-status">${statusLabel}</span>
        ${orderModeLabel ? `<span class="chip-meta">${orderModeLabel}</span>` : ""}
      </button>
    </div>
  `;
}

function itemOrderSource(item) {
  const customer = item.customerId ? customerById(item.customerId) : null;
  if (!customer) return "ONE-OFF";
  if (item.orderMode === "THEY_PUT_ORDER") {
    const channel = customer.orderChannel || customer.channelPreference || "portal";
    return `CUSTOMER ORDERS (${channelLabel(channel).toUpperCase()})`;
  }
  const channel = customer.orderChannel || customer.channelPreference || "portal";
  if (item.orderMode === "WE_GET_ORDER") return "WE GET ORDER";
  return channelLabel(channel).toUpperCase();
}

function scheduleOrderModeLabel(item) {
  if (item.orderMode === "THEY_PUT_ORDER") return "Customer orders";
  if (item.orderMode === "WE_GET_ORDER") return "We order";
  return "";
}

function itemExportType(item) {
  const baseKind = item.kind === "custom_oneoff" ? item.oneOffKind : item.kind;
  if (baseKind === "delivery") return "DELIVERY";
  if (baseKind === "pack") return "PACK";
  return "ORDER";
}

function buildSpokeExport(items) {
  const rows = items.map((item) => {
    const customer = item.customerId ? customerById(item.customerId) : null;
    const baseKind = item.kind === "custom_oneoff" ? item.oneOffKind : item.kind;
    const notes = [customer?.deliveryNotes, item.note].filter(Boolean).join(" • ");
    const repNameValue = baseKind === "delivery" ? repName(item.repId) : "";
    return [
      customer?.storeName || item.title,
      customer?.fullAddress || "",
      notes,
      itemExportType(item),
      itemOrderSource(item),
      item.date,
      repNameValue,
      customer?.customerId || "",
    ];
  });
  const header = [
    "stop_name",
    "address",
    "notes",
    "type",
    "order_source",
    "date",
    "rep_name",
    "customer_id",
  ];
  return [header, ...rows].map((row) => row.map(toCsvValue).join(",")).join("\n");
}

function renderDashboard() {
  const today = todayKey();
  const repFilter = elements.dashboardRepFilter.value;
  const searchTerm = elements.dashboardSearch.value || "";
  const toggles = state.settings.app.scheduleView.toggles;
  const statusFilter = state.settings.app.scheduleView.statusFilter;

  const items = buildAgendaItems({
    dateStart: today,
    dateEnd: today,
    toggles,
    statusFilter,
    repFilter,
    searchTerm,
  });

  const ordersPut = items.filter(
    (item) =>
      (item.kind === "expected_order" && item.orderMode === "THEY_PUT_ORDER") ||
      (item.kind === "custom_oneoff" &&
        item.oneOffKind === "expected_order" &&
        item.orderMode === "THEY_PUT_ORDER")
  );
  const ordersGet = items.filter(
    (item) =>
      (item.kind === "expected_order" && item.orderMode !== "THEY_PUT_ORDER") ||
      (item.kind === "custom_oneoff" &&
        item.oneOffKind === "expected_order" &&
        item.orderMode !== "THEY_PUT_ORDER")
  );
  const packs = items.filter(
    (item) => item.kind === "pack" || (item.kind === "custom_oneoff" && item.oneOffKind === "pack")
  );
  const deliveries = items.filter(
    (item) => item.kind === "delivery" || (item.kind === "custom_oneoff" && item.oneOffKind === "delivery")
  );

  elements.todayOrdersPutCount.textContent = ordersPut.length;
  elements.todayOrdersGetCount.textContent = ordersGet.length;
  elements.todayPacksCount.textContent = packs.length;
  elements.todayDeliveriesCount.textContent = deliveries.length;

  elements.todayOrdersPutList.innerHTML =
    ordersPut.map(agendaItemCard).join("") || "<p class=\"muted\">No customers placing orders today.</p>";
  elements.todayOrdersGetList.innerHTML =
    ordersGet.map(agendaItemCard).join("") || "<p class=\"muted\">No orders to take today.</p>";
  elements.todayPacksList.innerHTML =
    packs.map(agendaItemCard).join("") || "<p class=\"muted\">No packing today.</p>";
  elements.todayDeliveriesList.innerHTML =
    deliveries.map(agendaItemCard).join("") || "<p class=\"muted\">No deliveries today.</p>";

  attachAgendaItemActions(elements.todayOrdersPutList, ordersPut);
  attachAgendaItemActions(elements.todayOrdersGetList, ordersGet);
  attachAgendaItemActions(elements.todayPacksList, packs);
  attachAgendaItemActions(elements.todayDeliveriesList, deliveries);
}

function renderOrders() {
  elements.ordersList.innerHTML = state.orders
    .slice()
    .sort((a, b) => new Date(b.receivedAt) - new Date(a.receivedAt))
    .map((order) => {
      const customer = customerById(order.customerId);
      return `
        <div class="list-item">
          <header>
            <strong>${customer?.storeName || "Unknown"}</strong>
            <div class="badges">
              <span class="badge">${channelLabel(order.channel)}</span>
              <span class="badge">${statusLabels[order.status] || order.status}</span>
            </div>
          </header>
          <div class="muted">Received ${formatDateTime(order.receivedAt)}</div>
          <div class="muted">Pack due ${formatDate(order.packDueDate)} • Deliver ${formatDate(order.deliveryDueDate)}</div>
          <div class="muted">Rep: ${repName(order.assignedRepId)}</div>
          <div class="form-actions">
            <button class="secondary" data-order-id="${order.id}" data-action="edit">Edit</button>
            <button class="ghost" data-order-id="${order.id}" data-action="delete">Delete</button>
          </div>
        </div>
      `;
    })
    .join("") || "<p class=\"muted\">No orders yet.</p>";

  elements.ordersList.querySelectorAll("button[data-order-id]").forEach((button) => {
    button.addEventListener("click", async () => {
      const orderId = button.dataset.orderId;
      const action = button.dataset.action;
      const order = orderById(orderId);
      if (!order) return;
      if (action === "edit") {
        openOrderModal(order);
      }
      if (action === "delete") {
        if (!canWrite()) {
          showOfflineAlert();
          return;
        }
        if (!confirm("Delete this order and its tasks?")) return;
        await deleteItem("orders", orderId);
        state.orders = state.orders.filter((item) => item.id !== orderId);
        await removeTasksForOrder(orderId);
        renderAll();
      }
    });
  });
}

function customerMatchesFilters(customer) {
  const term = elements.customersSearch.value || "";
  const repFilter = elements.customersRepFilter.value;
  const modeFilter = elements.customersModeFilter.value;
  const freqFilter = elements.customersFrequencyFilter.value;
  const channelFilter = elements.customersChannelFilter.value;
  const scheduleFilter = elements.customersScheduleFilter.value;

  if (repFilter !== "all" && customer.assignedRepId !== repFilter) return false;
  if (modeFilter !== "all" && customer.schedule?.mode !== modeFilter) return false;
  if (freqFilter !== "all" && customer.schedule?.frequency !== freqFilter) return false;
  if (channelFilter !== "all") {
    const channel = customer.orderChannel || customer.channelPreference || "portal";
    if (channel !== channelFilter) return false;
  }
  if (scheduleFilter === "scheduled" && !customer.schedule?.frequency) return false;
  if (scheduleFilter === "unscheduled" && customer.schedule?.frequency) return false;

  const combined = [
    customer.storeName,
    customer.contactName,
    customer.fullAddress,
    customer.suburb1,
    customer.postcode1,
    customer.phone,
    customer.email,
  ]
    .filter(Boolean)
    .join(" ");
  return matchesSearch(combined, term);
}

function renderCustomers() {
  elements.customersList.innerHTML = state.customers
    .slice()
    .sort((a, b) => a.storeName.localeCompare(b.storeName))
    .filter((customer) => customerMatchesFilters(customer))
    .map((customer) => {
      const repLabel = repName(customer.assignedRepId);
      const initials = customer.storeName
        .split(" ")
        .map((part) => part[0])
        .slice(0, 2)
        .join("")
        .toUpperCase();
      return `
      <button class="customer-row" data-customer-id="${customer.id}" data-action="edit">
        <div class="customer-avatar">${initials || "?"}</div>
        <div class="customer-meta">
          <div class="customer-name">${customer.storeName}</div>
          <div class="muted">${customer.fullAddress}</div>
        </div>
        <div class="customer-actions">
          <span class="pill">${repLabel}</span>
          <span class="chevron">›</span>
        </div>
      </button>
    `;
    })
    .join("") || "<p class=\"muted\">No customers yet.</p>";

  elements.customersList.querySelectorAll("button[data-customer-id]").forEach((button) => {
    button.addEventListener("click", () => {
      const customerId = button.dataset.customerId;
      const customer = customerById(customerId);
      if (customer) openCustomerModal(customer);
    });
  });
}

function renderReps() {
  elements.repsList.innerHTML = state.reps
    .slice()
    .sort((a, b) => a.name.localeCompare(b.name))
    .map((rep) => `
      <div class="list-item">
        <header>
          <strong>${rep.name}</strong>
          <div class="badges">
            <span class="badge">${rep.active !== false ? "Active" : "Inactive"}</span>
          </div>
        </header>
        <div class="muted">${rep.phone || "No phone"} • ${rep.vehicle || "No vehicle"}</div>
        <div class="form-actions">
          <button class="secondary" data-rep-id="${rep.id}" data-action="edit">Edit</button>
        </div>
      </div>
    `)
    .join("") || "<p class=\"muted\">No reps yet.</p>";

  elements.repsList.querySelectorAll("button[data-rep-id]").forEach((button) => {
    button.addEventListener("click", () => {
      const repId = button.dataset.repId;
      const rep = state.reps.find((item) => item.id === repId);
      if (rep) openRepModal(rep);
    });
  });
}

function computeNextExpected(customer) {
  if (!customer.cadenceType) {
    return null;
  }
  const today = todayKey();
  const lastOrder = state.orders
    .filter((order) => order.customerId === customer.id)
    .sort((a, b) => new Date(b.receivedAt) - new Date(a.receivedAt))[0];

  const lastDateKey = lastOrder ? todayKeyFromISO(lastOrder.receivedAt) : today;

  if (customer.cadenceType === "weekly") {
    return nextWeeklyDate(customer.cadenceDayOfWeek, today);
  }

  if (customer.cadenceType === "fortnightly") {
    return customer.cadenceNextDueDate || addDays(lastDateKey, 14);
  }

  if (customer.cadenceType === "twice_weekly") {
    return nextTwiceWeeklyDate(customer.cadenceDaysOfWeek || [], today);
  }

  if (customer.cadenceType === "custom") {
    if (customer.cadenceEveryNDays) {
      return addDays(lastDateKey, Number(customer.cadenceEveryNDays));
    }
    return nextWeeklyDate(customer.cadenceDayOfWeek, today);
  }

  return today;
}

function nextWeeklyDate(dayOfWeek, startKey) {
  if (!dayOfWeek) return addDays(startKey, 7);
  const map = {
    Monday: 1,
    Tuesday: 2,
    Wednesday: 3,
    Thursday: 4,
    Friday: 5,
    Saturday: 6,
    Sunday: 0,
  };
  const target = map[dayOfWeek] ?? 1;
  const current = dayIndexFromDateKey(startKey);
  let diff = target - current;
  if (diff <= 0) diff += 7;
  return addDays(startKey, diff);
}

function nextTwiceWeeklyDate(days, startKey) {
  const sorted = days
    .map((day) => ({
      day,
      key: nextWeeklyDate(day, startKey),
    }))
    .sort((a, b) => a.key.localeCompare(b.key));
  return sorted[0]?.key || addDays(startKey, 3);
}

function renderExpected() {
  const today = todayKey();
  const upcomingKeys = new Set(dateRange(today, 7));
  const expected = state.customers
    .map((customer) => ({
      customer,
      nextDate: computeNextExpected(customer),
      lastOrder: state.orders
        .filter((order) => order.customerId === customer.id)
        .sort((a, b) => new Date(b.receivedAt) - new Date(a.receivedAt))[0],
    }))
    .filter((item) => item.nextDate && upcomingKeys.has(item.nextDate));

  elements.expectedList.innerHTML = expected
    .map((item) => `
      <div class="list-item">
        <header>
          <strong>${item.customer.storeName}</strong>
          <div class="badges">
            <span class="badge">${repName(item.customer.assignedRepId)}</span>
            <span class="badge">${formatDate(item.nextDate)}</span>
          </div>
        </header>
        <div class="muted">Last order: ${item.lastOrder ? formatDateTime(item.lastOrder.receivedAt) : "None"}</div>
        <div class="form-actions">
          <button class="secondary" data-customer-id="${item.customer.id}" data-action="placeholder">Create placeholder</button>
          <button class="ghost" data-customer-id="${item.customer.id}" data-action="log">Log order</button>
        </div>
      </div>
    `)
    .join("") || "<p class=\"muted\">No expected orders in the next 7 days.</p>";

  elements.expectedList.querySelectorAll("button[data-customer-id]").forEach((button) => {
    button.addEventListener("click", () => {
      const customer = customerById(button.dataset.customerId);
      if (!customer) return;
      if (button.dataset.action === "placeholder") {
        createPlaceholderOrder(customer);
      } else {
        openOrderModal({ customerId: customer.id });
      }
    });
  });
}

async function createPlaceholderOrder(customer) {
  if (!canWrite()) {
    showOfflineAlert();
    return;
  }
  const order = buildOrderFromForm({
    customerId: customer.id,
    channel: customer.orderChannel || customer.channelPreference || "portal",
    receivedAt: toISO(),
    orderLines: [],
  });
  await saveOrder(order, true);
  renderAll();
}

function renderSchedule() {
  const view = state.settings.app.scheduleView;
  const repFilter = elements.scheduleRepFilter.value;
  const toggles = view.toggles;
  const statusFilter = view.statusFilter;
  const searchTerm = view.searchTerm || "";
  const anchorDate = view.anchorDate || todayKey();
  const viewMode = view.viewMode || "week";
  const activeDayKey = view.activeDayKey || anchorDate;

  if (elements.scheduleToolbar) {
    elements.scheduleToolbar.dataset.view = viewMode;
  }
  if (elements.scheduleDaySelect) {
    elements.scheduleDaySelect.disabled = viewMode !== "week";
  }
  if (elements.scheduleDayPrev) {
    elements.scheduleDayPrev.disabled = viewMode !== "week";
  }
  if (elements.scheduleDayNext) {
    elements.scheduleDayNext.disabled = viewMode !== "week";
  }

  elements.scheduleRangeLabel.textContent = formatRangeLabel(viewMode, anchorDate);
  elements.scheduleDatePicker.value = anchorDate;

  const baseKindForItem = (item) =>
    item.kind === "custom_oneoff" ? item.oneOffKind : item.kind;
  const sortByTypeThenTitle = (a, b) => {
    const diff = agendaSortOrder(a) - agendaSortOrder(b);
    if (diff !== 0) return diff;
    return a.title.localeCompare(b.title);
  };

  if (viewMode === "day") {
    if (elements.scheduleDaySelect) {
      elements.scheduleDaySelect.innerHTML = `<option value="${anchorDate}">${formatDate(anchorDate)}</option>`;
      elements.scheduleDaySelect.value = anchorDate;
    }
    const items = getItemsForRange({
      dateStart: anchorDate,
      dateEnd: anchorDate,
      toggles,
      statusFilter,
      repFilter,
      searchTerm,
    }).sort(sortByTypeThenTitle);
    const deliveries = items.filter((item) => baseKindForItem(item) === "delivery");
    const packs = items.filter((item) => baseKindForItem(item) === "pack");
    const orders = items.filter((item) => baseKindForItem(item) === "expected_order");
    elements.scheduleList.innerHTML = `
      <div class="schedule-day-view">
        <div class="schedule-lane">
          <div class="lane-header">
            <span class="type-label delivery">Deliveries</span>
            <span class="pill">${deliveries.length}</span>
          </div>
          <div class="lane-items">
            ${deliveries.map((item) => scheduleItemDetailCard(item)).join("") || "<p class=\"muted\">No deliveries.</p>"}
          </div>
        </div>
        <div class="schedule-lane">
          <div class="lane-header">
            <span class="type-label pack">Packing</span>
            <span class="pill">${packs.length}</span>
          </div>
          <div class="lane-items">
            ${packs.map((item) => scheduleItemDetailCard(item)).join("") || "<p class=\"muted\">No packing.</p>"}
          </div>
        </div>
        <div class="schedule-lane">
          <div class="lane-header">
            <span class="type-label order">Expected Orders</span>
            <span class="pill">${orders.length}</span>
          </div>
          <div class="lane-items">
            ${orders.map((item) => scheduleItemDetailCard(item)).join("") || "<p class=\"muted\">No expected orders.</p>"}
          </div>
        </div>
      </div>
    `;
    attachScheduleSelectionHandlers(elements.scheduleList, items);
    attachScheduleDetailActions(elements.scheduleList, items);
    return;
  }

  if (viewMode === "week") {
    const weekStart = startOfWeekMonday(anchorDate);
    const weekEnd = addDays(weekStart, 6);
    const items = getItemsForRange({
      dateStart: weekStart,
      dateEnd: weekEnd,
      toggles,
      statusFilter,
      repFilter,
      searchTerm,
    }).sort(sortByTypeThenTitle);
    const grouped = items.reduce((acc, item) => {
      acc[item.date] = acc[item.date] || [];
      acc[item.date].push(item);
      return acc;
    }, {});
    const days = dateRange(weekStart, 7);
    const resolvedActiveDay = days.includes(activeDayKey) ? activeDayKey : anchorDate;
    if (resolvedActiveDay !== view.activeDayKey) {
      state.settings.app.scheduleView.activeDayKey = resolvedActiveDay;
    }
    if (elements.scheduleDaySelect) {
      elements.scheduleDaySelect.innerHTML = days
        .map((dateKey) => `<option value="${dateKey}">${formatDate(dateKey)}</option>`)
        .join("");
      elements.scheduleDaySelect.value = resolvedActiveDay;
    }
    elements.scheduleList.innerHTML = `
      <div class="schedule-week-scroll">
        <div class="schedule-week-view">
          ${days
            .map((dateKey) => {
              const dayItems = (grouped[dateKey] || []).sort(sortByTypeThenTitle);
              const isToday = dateKey === todayKey();
              const isActive = dateKey === resolvedActiveDay;
              return `
                <div class="week-day ${isToday ? "today" : ""} ${isActive ? "is-active" : "is-inactive"}" data-date="${dateKey}">
                  <div class="week-day-header">
                    <span>${formatDate(dateKey)}</span>
                    ${isToday ? '<span class="pill">Today</span>' : ""}
                  </div>
                  <div class="week-day-items">
                    ${dayItems.map((item) => `<div class="week-item" data-item-id="${item.selectionId}">${scheduleItemChip(item)}</div>`).join("") || "<span class=\"muted\">—</span>"}
                  </div>
                </div>
              `;
            })
            .join("")}
        </div>
      </div>
    `;
    attachScheduleSelectionHandlers(elements.scheduleList, items);
    attachScheduleChipActions(elements.scheduleList, items);
    return;
  }

  if (elements.scheduleDaySelect) {
    elements.scheduleDaySelect.innerHTML = `<option value="${anchorDate}">${formatDate(anchorDate)}</option>`;
    elements.scheduleDaySelect.value = anchorDate;
  }
  const { gridStart, gridEnd } = monthGridRange(anchorDate);
  const totalDays =
    Math.floor((dateKeyToDate(gridEnd) - dateKeyToDate(gridStart)) / (1000 * 60 * 60 * 24)) + 1;
  const gridDays = dateRange(gridStart, totalDays);
  const items = getItemsForRange({
    dateStart: gridStart,
    dateEnd: gridEnd,
    toggles,
    statusFilter,
    repFilter,
    searchTerm,
  }).sort(sortByTypeThenTitle);
  const grouped = items.reduce((acc, item) => {
    acc[item.date] = acc[item.date] || [];
    acc[item.date].push(item);
    return acc;
  }, {});
  const weekLabels = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];
  elements.scheduleList.innerHTML = `
    <div class="schedule-month-view">
      <div class="month-grid-header">
        ${weekLabels.map((label) => `<div class="month-header-cell">${label}</div>`).join("")}
      </div>
      <div class="month-grid">
        ${gridDays
          .map((dateKey) => {
            const dayItems = (grouped[dateKey] || []).sort(sortByTypeThenTitle);
            const isToday = dateKey === todayKey();
            const inMonth = dateKey.startsWith(anchorDate.slice(0, 7));
            const visibleItems = dayItems.slice(0, 3);
            const extraCount = dayItems.length - visibleItems.length;
            return `
              <div class="month-cell ${isToday ? "today" : ""} ${inMonth ? "" : "outside"}" data-date="${dateKey}">
                <div class="month-cell-header">
                  <span>${dateKey.split("-")[2]}</span>
                </div>
                <div class="month-cell-items">
                  ${visibleItems
                    .map(
                      (item) =>
                        `<div class="month-item" data-item-id="${item.selectionId}">${scheduleItemChip(item)}</div>`
                    )
                    .join("")}
                  ${extraCount > 0 ? `<button class="more-button" data-action="more">+${extraCount} more</button>` : ""}
                </div>
              </div>
            `;
          })
          .join("")}
      </div>
    </div>
  `;
  attachScheduleSelectionHandlers(elements.scheduleList, items);
  attachScheduleChipActions(elements.scheduleList, items, true);
}

function attachScheduleDetailActions(container, items) {
  const itemMap = new Map(items.map((item) => [item.selectionId, item]));
  container.querySelectorAll("[data-action]").forEach((button) => {
    const itemElement = button.closest("[data-item-id]");
    if (!itemElement) return;
    const item = itemMap.get(itemElement.dataset.itemId);
    if (!item) return;
    button.addEventListener("click", async () => {
      const action = button.dataset.action;
      if (action === "toggle-details") {
        const card = button.closest(".schedule-detail-card");
        if (!card) return;
        const isCollapsed = card.classList.toggle("is-collapsed");
        button.textContent = isCollapsed ? "Expand" : "Collapse";
        button.setAttribute("aria-expanded", String(!isCollapsed));
        return;
      }
      await handleAgendaAction(item, action);
    });
  });
}

function attachScheduleChipActions(container, items, isMonth = false) {
  const itemMap = new Map(items.map((item) => [item.selectionId, item]));
  container.querySelectorAll(".schedule-chip").forEach((chip) => {
    const itemElement = chip.closest("[data-item-id]");
    const item = itemElement ? itemMap.get(itemElement.dataset.itemId) : null;
    if (!item) return;
    chip.addEventListener("click", (event) => {
      event.stopPropagation();
      openSchedulePopover(item, chip);
    });
  });
  if (isMonth) {
    container.querySelectorAll(".more-button").forEach((button) => {
      button.addEventListener("click", (event) => {
        event.stopPropagation();
        const cell = button.closest(".month-cell");
        if (!cell) return;
        const dateKey = cell.dataset.date;
        openDayModal(dateKey, items.filter((item) => item.date === dateKey));
      });
    });
    container.querySelectorAll(".month-cell").forEach((cell) => {
      cell.addEventListener("click", () => {
        const dateKey = cell.dataset.date;
        openDayModal(dateKey, items.filter((item) => item.date === dateKey));
      });
    });
  }
}

function openDayModal(dateKey, items) {
  const content = `
    <h2>${formatDate(dateKey)}</h2>
    <div class="schedule-modal-list">
      ${items.length ? items.map((item) => scheduleItemDetailCard(item)).join("") : "<p class=\"muted\">No items.</p>"}
    </div>
  `;
  showModal(content);
  attachScheduleDetailActions(document.getElementById("modalBody"), items);
}

function openSchedulePopover(item, anchor) {
  const popover = document.getElementById("schedulePopover");
  if (!popover) return;
  popover.innerHTML = `
    <div class="schedule-popover-card" data-item-id="${item.selectionId}">
      <button class="popover-close" type="button" data-action="close">×</button>
      ${scheduleItemDetailCard(item)}
    </div>
  `;
  const rect = anchor.getBoundingClientRect();
  popover.style.top = `${window.scrollY + rect.bottom + 8}px`;
  popover.style.left = `${window.scrollX + rect.left}px`;
  popover.classList.add("active");
  const closePopover = () => {
    popover.classList.remove("active");
  };
  popover.querySelector("[data-action='close']").addEventListener("click", closePopover);
  document.addEventListener(
    "click",
    (event) => {
      if (!popover.contains(event.target) && !anchor.contains(event.target)) {
        closePopover();
      }
    },
    { once: true }
  );
  attachScheduleDetailActions(popover, [item]);
}

async function handleAgendaAction(item, action) {
  if (!canWrite() && ["done", "skip", "undo", "delete", "bump-back"].includes(action)) {
    showOfflineAlert();
    return;
  }
  if (action === "open") {
    const customer = item.customerId ? customerById(item.customerId) : null;
    if (customer) openCustomerModal(customer);
    return;
  }
  if (action === "done") {
    await setScheduleEventStatus(item, "DONE");
    showSnackbar("Marked done.", () => setScheduleEventStatus(item, "PENDING"));
    return;
  }
  if (action === "skip") {
    openSkipModal(item);
    return;
  }
  if (action === "undo") {
    await setScheduleEventStatus(item, "PENDING");
    return;
  }
  if (action === "bump-back" && item.isOneOff) {
    await bumpOneOffItemDate(item, -1);
    return;
  }
  if (action === "delete" && item.isOneOff) {
    await deleteOneOffItem(item.sourceId);
  }
}

function attachAgendaItemActions(container, items) {
  const itemMap = new Map(items.map((item) => [item.selectionId, item]));
  container.querySelectorAll(".list-item").forEach((itemElement) => {
    const itemId = itemElement.dataset.itemId;
    const item = itemMap.get(itemId);
    if (!item) return;
    itemElement.querySelectorAll("button[data-action], input[data-action]").forEach((button) => {
      button.addEventListener("click", async () => {
        const action = button.dataset.action;
        await handleAgendaAction(item, action);
      });
    });
  });
}

async function bumpOneOffItemDate(item, offsetDays) {
  if (!canWrite()) {
    showOfflineAlert();
    return;
  }
  const record = state.oneOffItems.find((entry) => entry.id === item.sourceId);
  if (!record) return;
  const previousDate = record.date;
  const nextDate = addDays(previousDate, offsetDays);
  record.date = nextDate;
  await put("one_off_items", record);
  state.oneOffItems = state.oneOffItems.map((entry) => (entry.id === record.id ? record : entry));
  renderAll();
  showSnackbar(`Bumped back to ${formatDate(nextDate)}.`, async () => {
    record.date = previousDate;
    await put("one_off_items", record);
    state.oneOffItems = state.oneOffItems.map((entry) =>
      entry.id === record.id ? record : entry
    );
    renderAll();
  });
}

async function deleteOneOffItem(oneOffId) {
  if (!canWrite()) {
    showOfflineAlert();
    return;
  }
  const item = state.oneOffItems.find((entry) => entry.id === oneOffId);
  if (!item) return;
  const previous = { ...item };
  item.isDeleted = true;
  await put("one_off_items", item);
  state.oneOffItems = state.oneOffItems.map((entry) => (entry.id === item.id ? item : entry));
  renderAll();
  showSnackbar("One-off item deleted.", async () => {
    item.isDeleted = previous.isDeleted;
    await put("one_off_items", item);
    state.oneOffItems = state.oneOffItems.map((entry) => (entry.id === item.id ? item : entry));
    renderAll();
  });
}

function openSkipModal(item) {
  showModal(`
    <h2>Skip item</h2>
    <form id="skipForm" class="form">
      <label><input type="radio" name="reason" value="ENOUGH_STOCK" required /> Enough stock</label>
      <label><input type="radio" name="reason" value="NO_ANSWER" /> Customer didn’t answer</label>
      <label><input type="radio" name="reason" value="UNSPECIFIED" /> Unspecified</label>
      <label><input type="radio" name="reason" value="CUSTOM" /> Custom reason</label>
      <label>Custom reason
        <input name="customReason" placeholder="Short reason" />
      </label>
      <div class="form-actions">
        <button class="secondary" type="button" id="cancelSkip">Cancel</button>
        <button class="primary" type="submit">Save</button>
      </div>
    </form>
  `);
  disableFormIfOffline("skipForm");

  document.getElementById("cancelSkip").addEventListener("click", closeModal);
  document.getElementById("skipForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    const formData = new FormData(event.target);
    const reason = formData.get("reason");
    const customReason = formData.get("customReason").trim();
    if (reason === "CUSTOM" && !customReason) {
      alert("Please enter a custom reason.");
      return;
    }
    await setScheduleEventStatus(item, "SKIPPED", {
      skippedReason: reason,
      skippedReasonText: reason === "CUSTOM" ? customReason : null,
    });
    closeModal();
  });
}

function openOneOffModal() {
  const customerOptions = [
    '<option value="">No customer</option>',
    ...state.customers.map((customer) => `<option value="${customer.id}">${customer.storeName}</option>`),
  ].join("");
  const repOptions = [
    '<option value="">Use customer rep</option>',
    ...state.reps.map((rep) => `<option value="${rep.id}">${rep.name}</option>`),
  ].join("");
  showModal(`
    <h2>Add one-off item</h2>
    <form id="oneOffForm" class="form">
      <label>Type ${helpIcon("Choose the one-off type. Example: One-off Delivery for a special drop this week.")}
        <select name="kind" id="oneOffKind">
          <option value="expected_order">One-off Expected Order</option>
          <option value="pack">One-off Pack</option>
          <option value="delivery">One-off Delivery</option>
        </select>
      </label>
      <label>Date ${helpIcon("Pick the date this one-off should appear on the agenda.")}
        <input name="date" type="date" value="${todayKey()}" />
      </label>
      <label>Customer ${helpIcon("Choose a customer, or leave blank for an ad-hoc order or pack.")}
        <select name="customerId" id="oneOffCustomer">
          ${customerOptions}
        </select>
      </label>
      <label>Assigned rep ${helpIcon("Optional. Leave blank to use the customer’s assigned rep.")}
        <select name="repId">
          ${repOptions}
        </select>
      </label>
      <label>Optional note ${helpIcon("Optional short note shown on the agenda.")}
        <input name="note" maxlength="120" placeholder="Optional short note" />
      </label>
      <div id="oneOffDeliveryHint" class="muted"></div>
      <div class="form-actions">
        <button class="secondary" type="button" id="cancelOneOff">Cancel</button>
        <button class="primary" type="submit">Save</button>
      </div>
    </form>
  `);
  disableFormIfOffline("oneOffForm");

  const kindSelect = document.getElementById("oneOffKind");
  const customerSelect = document.getElementById("oneOffCustomer");
  const repSelect = document.querySelector("select[name='repId']");
  const deliveryHint = document.getElementById("oneOffDeliveryHint");

  const updateDeliveryHint = () => {
    const kind = kindSelect.value;
    const customer = customerById(customerSelect.value);
    if (customer && !repSelect.value) {
      repSelect.value = customer.assignedRepId || "";
    }
    if (kind === "delivery" && !customer) {
      deliveryHint.textContent = "Delivery one-offs require a customer so we can use their saved address.";
    } else if (kind === "delivery" && customer) {
      deliveryHint.textContent = `Delivery will use: ${customer.fullAddress}`;
    } else {
      deliveryHint.textContent = "";
    }
  };

  kindSelect.addEventListener("change", updateDeliveryHint);
  customerSelect.addEventListener("change", updateDeliveryHint);
  updateDeliveryHint();

  document.getElementById("cancelOneOff").addEventListener("click", closeModal);
  document.getElementById("oneOffForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    if (!canWrite()) {
      showOfflineAlert();
      return;
    }
    const formData = new FormData(event.target);
    const kind = formData.get("kind");
    const date = formData.get("date");
    const customerId = formData.get("customerId") || null;
    const repId = formData.get("repId") || null;
    const note = formData.get("note").trim();
    if (kind === "delivery" && !customerId) {
      alert("Delivery one-offs require a customer so we can use the saved address.");
      return;
    }
    if (!date) {
      alert("Please choose a date.");
      return;
    }
    const record = {
      id: uuid(),
      kind,
      date,
      customerId,
      repId,
      note,
      isDeleted: false,
    };
    await put("one_off_items", record);
    state.oneOffItems.push(record);
    closeModal();
    renderAll();
  });
}

async function setScheduleEventStatus(item, status, details = {}) {
  if (!canWrite()) {
    showOfflineAlert();
    return;
  }
  const key = scheduleEventKeyForItem(item);
  const existing = state.scheduleEvents.find((event) => scheduleEventLookupKey(event) === key);
  if (status === "PENDING") {
    if (existing) {
      try {
        await syncDeleteScheduleEvent(existing.id);
        await deleteItem("schedule_events", existing.id);
        updateConnectionStatus({ status: "Online/Synced", canWrite: true });
      } catch (error) {
        console.error("Supabase delete failed", error);
        handleSupabaseError(error, { context: "Supabase delete failed", alertOnOffline: true });
        return;
      }
      state.scheduleEvents = state.scheduleEvents.filter((event) => event.id !== existing.id);
      renderAll();
    }
    return;
  }

  const payload = {
    id: existing?.id || createCustomerId(),
    customerId: item.customerId || null,
    kind: item.kind,
    date: item.date,
    runIndex: item.runIndex ?? null,
    sourceId: item.sourceId || null,
    status,
    skippedReason: details.skippedReason || null,
    skippedReasonText: details.skippedReasonText || null,
    completedAt: status === "DONE" ? toISO() : null,
    note: item.note || null,
  };

  try {
    await syncUpsertScheduleEvent(payload);
    await put("schedule_events", payload);
    updateConnectionStatus({ status: "Online/Synced", canWrite: true });
    setCloudStatus("synced", toISO());
  } catch (error) {
    console.error("Supabase upsert failed", error);
    handleSupabaseError(error, { context: "Supabase upsert failed", alertOnOffline: true });
    return;
  }
  if (existing) {
    state.scheduleEvents = state.scheduleEvents.map((event) => (event.id === payload.id ? payload : event));
  } else {
    state.scheduleEvents.push(payload);
  }
  renderAll();
}

function renderExport() {
  const today = todayKey();
  elements.exportStart.value = today;
  elements.exportEnd.value = today;
  renderMappingPresets();
}

function renderMappingPresets() {
  const presets = state.settings?.exportPresets ?? state.settings?.app?.exportPresets ?? [];
  if (!presets.length) {
    elements.mappingPresetSelect.innerHTML = "<option value=\"\">No presets yet</option>";
    elements.mappingEditor.innerHTML = "<p class=\"muted\">No presets yet.</p>";
    return;
  }
  elements.mappingPresetSelect.innerHTML = presets
    .map((preset) => `<option value="${preset.id}">${preset.name}</option>`)
    .join("");
  elements.mappingPresetSelect.value =
    state.settings?.app?.activeExportPresetId || presets[0]?.id;
  renderMappingEditor();
}

function renderMappingEditor() {
  const preset = getActivePreset();
  if (!preset) return;
  elements.mappingEditor.innerHTML = preset.columns
    .map((column, index) => `
      <div class="mapping-row">
        <input data-column-index="${index}" data-field="header" value="${column.header}" />
        <select data-column-index="${index}" data-field="field">
          ${fieldOptions
            .map(
              (option) =>
                `<option value="${option.value}" ${option.value === column.field ? "selected" : ""}>${option.label}</option>`
            )
            .join("")}
        </select>
        <button class="ghost" data-action="remove" data-column-index="${index}">Remove</button>
      </div>
    `)
    .join("");

  elements.mappingEditor.innerHTML += `
    <button id="addColumnBtn" class="secondary">Add column</button>
  `;

  elements.mappingEditor.querySelectorAll("button[data-action='remove']").forEach((button) => {
    button.addEventListener("click", () => {
      const index = Number(button.dataset.columnIndex);
      preset.columns.splice(index, 1);
      renderMappingEditor();
    });
  });

  const addBtn = document.getElementById("addColumnBtn");
  if (addBtn) {
    addBtn.addEventListener("click", () => {
      preset.columns.push({ header: "new_column", field: "customer.storeName" });
      renderMappingEditor();
    });
  }

  elements.mappingEditor.querySelectorAll("input, select").forEach((field) => {
    field.addEventListener("change", () => {
      const index = Number(field.dataset.columnIndex);
      const key = field.dataset.field;
      preset.columns[index][key] = field.value;
    });
  });
}

function getActivePreset() {
  return state.settings.app.exportPresets.find(
    (preset) => preset.id === state.settings.app.activeExportPresetId
  );
}

function mapOrderTerms(value) {
  if (!value) return null;
  const normalized = value.trim().toLowerCase();
  return orderTermsMap[normalized] || null;
}

function selectImportHeaders(headers) {
  const normalized = headers.map((header) => normalizeHeader(header));
  const getIndex = (options) =>
    options
      .map((option) => normalized.indexOf(normalizeHeader(option)))
      .find((index) => index !== -1) ?? -1;

  return {
    customerId: getIndex(["Customer ID"]),
    storeName: getIndex([
      "Store Name",
      "Customer Name",
      "Customer Name (If you change the customer name in any way, it will create another customer)",
    ]),
    phone: getIndex(["Phone", "Phone 1"]),
    email: getIndex(["Email", "Primary Email"]),
    address1: getIndex(["Delivery Address 1", "Delivery Address"]),
    suburb1: getIndex(["Delivery Suburb 1", "Delivery Suburb"]),
    state1: getIndex(["Delivery State 1", "Delivery State"]),
    postcode1: getIndex(["Delivery Postcode 1", "Delivery Postcode"]),
    deliveryNotes: getIndex(["Delivery Notes"]),
    deliveryTerms: getIndex(["Delivery terms"]),
    orderTermsLabel: getIndex(["Order terms"]),
    aovPrimary: getIndex(["Average Order Value (AUD)"]),
    aovLegacy: getIndex([
      "Average Order Value Lasdt 365 days (30/7/25",
      "Average Order Value Lasdt 365 days (30/7/25)",
    ]),
  };
}

function buildFullAddress(row, headerMap) {
  const parts = [
    row[headerMap.address1] || "",
    row[headerMap.suburb1] || "",
    row[headerMap.state1] || "",
    row[headerMap.postcode1] || "",
  ]
    .map((value) => String(value || "").trim())
    .filter(Boolean);
  return parts.join(", ");
}

function getImportAovMode() {
  return document.querySelector("input[name='importAovMode']:checked")?.value || "blank";
}

function buildImportPreview(rows, headers) {
  const headerMap = selectImportHeaders(headers);
  importState.headers = headers;
  importState.rows = rows;
  importState.headerMap = headerMap;
  const mappedRows = rows.map((row, index) => {
    const customerId = headerMap.customerId !== -1 ? String(row[headerMap.customerId] || "").trim() : "";
    const storeName = headerMap.storeName !== -1 ? String(row[headerMap.storeName] || "").trim() : "";
    const orderTerms = headerMap.orderTermsLabel !== -1 ? String(row[headerMap.orderTermsLabel] || "").trim() : "";
    const address = buildFullAddress(row, headerMap);
    const mappedOrderTerms = mapOrderTerms(orderTerms);
    const valid = Boolean((customerId || storeName) && address && mappedOrderTerms);
    return {
      index,
      customerId,
      storeName,
      orderTerms,
      mappedOrderTerms,
      address,
      status: valid ? "ok" : "needs",
    };
  });

  importState.mapped = mappedRows;
  renderImportPreview();
}

function getAovForRow(row, aovMode) {
  if (aovMode === "single") {
    const value = Number(elements.importAovValue.value || 0);
    return Number.isNaN(value) ? null : value;
  }
  if (aovMode === "column") {
    if (!elements.importAovColumn.value) {
      return null;
    }
    const columnIndex = Number(elements.importAovColumn.value);
    const raw = row[columnIndex];
    const value = Number(String(raw || "").trim());
    return Number.isNaN(value) ? null : value;
  }
  return null;
}

function renderImportPreview() {
  if (!importState.mapped.length) {
    elements.importPreview.innerHTML = "<p class=\"muted\">No preview yet.</p>";
    return;
  }

  const repNameLabel = elements.importDefaultRep.value
    ? repName(elements.importDefaultRep.value)
    : "Select rep";
  const aovMode = getImportAovMode();

  const tableRows = importState.mapped
    .slice(0, 20)
    .map(
      (row) => `
      <tr>
        <td>${row.index + 1}</td>
        <td>${row.customerId || "—"}</td>
        <td>${row.storeName || "—"}</td>
        <td>${row.address || "—"}</td>
        <td>
          <select data-import-row="${row.index}">
            <option value="">Select...</option>
            ${orderTermsOptions
              .map(
                (option) =>
                  `<option value="${option.value}" ${option.value === row.mappedOrderTerms ? "selected" : ""}>${option.label}</option>`
              )
              .join("")}
          </select>
        </td>
        <td>${repNameLabel || "Select rep"}</td>
        <td>${getAovForRow(importState.rows[row.index], aovMode) ?? "—"}</td>
        <td><span class="status-badge ${row.status === "ok" ? "ok" : ""}">${row.status === "ok" ? "Ready" : "Needs attention"}</span></td>
      </tr>
    `
    )
    .join("");

  elements.importPreview.innerHTML = `
    <table class="preview-table">
      <thead>
        <tr>
          <th>#</th>
          <th>Customer ID</th>
          <th>Store name</th>
          <th>Address</th>
          <th>Order terms</th>
          <th>Rep</th>
          <th>AOV (AUD)</th>
          <th>Status</th>
        </tr>
      </thead>
      <tbody>${tableRows}</tbody>
    </table>
  `;

  elements.importPreview.querySelectorAll("select[data-import-row]").forEach((select) => {
    select.addEventListener("change", () => {
      const rowIndex = Number(select.dataset.importRow);
      const row = importState.mapped.find((item) => item.index === rowIndex);
      if (!row) return;
      row.mappedOrderTerms = select.value || null;
      row.status = row.mappedOrderTerms && (row.customerId || row.storeName) && row.address ? "ok" : "needs";
      renderImportPreview();
    });
  });
}

function toCsvValue(value) {
  if (value === null || value === undefined) return "";
  const stringValue = String(value);
  if (stringValue.includes(",") || stringValue.includes("\n") || stringValue.includes("\"")) {
    return `"${stringValue.replace(/\"/g, "\"\"")}"`;
  }
  return stringValue;
}

async function handleImportPreview() {
  const file = elements.importFile.files[0];
  if (!file) {
    alert("Select a CSV file first.");
    return;
  }
  const text = await file.text();
  const rows = parseCsv(text);
  if (rows.length < 2) {
    alert("CSV appears to be empty.");
    return;
  }
  const headers = rows[0].map((header) => header.trim());
  const dataRows = rows.slice(1);
  const headerMap = selectImportHeaders(headers);
  if (headerMap.orderTermsLabel === -1 || headerMap.address1 === -1) {
    alert("Missing required headers. Ensure Order terms and Delivery Address columns exist.");
    return;
  }
  importState.headers = headers;
  importState.rows = dataRows;
  const numericColumns = detectNumericColumns(headers, dataRows.slice(0, 50));
  elements.importAovColumn.innerHTML = [
    "<option value=\"\">Select column</option>",
    ...numericColumns.map((column) => `<option value="${column.index}">${column.header}</option>`),
  ].join("");
  const preferredAovIndex = headerMap.aovPrimary !== -1 ? headerMap.aovPrimary : headerMap.aovLegacy;
  if (preferredAovIndex !== -1) {
    elements.importAovColumn.value = String(preferredAovIndex);
  }
  buildImportPreview(dataRows, headers);
  elements.importStatus.textContent = `${dataRows.length} rows detected.`;
}

function getExistingCustomerByExternalId(externalId) {
  if (!externalId) return null;
  return (
    state.customers.find((customer) => customer.customerId === externalId) ||
    state.customers.find((customer) => customer.externalCustomerId === externalId) ||
    state.customers.find((customer) => customer.id === externalId)
  );
}

function buildCustomerFromRow(row, rowIndex) {
  const headerMap = importState.headerMap;
  const customerId = headerMap.customerId !== -1 ? String(row[headerMap.customerId] || "").trim() : "";
  const storeName = headerMap.storeName !== -1 ? String(row[headerMap.storeName] || "").trim() : "";
  const orderTermsLabel = headerMap.orderTermsLabel !== -1 ? String(row[headerMap.orderTermsLabel] || "").trim() : "";
  const mappedOrderTerms = importState.mapped.find((item) => item.index === rowIndex)?.mappedOrderTerms;
  const fullAddress = buildFullAddress(row, headerMap);
  const address1 = headerMap.address1 !== -1 ? String(row[headerMap.address1] || "").trim() : "";
  const suburb = headerMap.suburb1 !== -1 ? String(row[headerMap.suburb1] || "").trim() : "";
  const stateValue = headerMap.state1 !== -1 ? String(row[headerMap.state1] || "").trim() : "";
  const postcode = headerMap.postcode1 !== -1 ? String(row[headerMap.postcode1] || "").trim() : "";
  const deliveryNotes = headerMap.deliveryNotes !== -1 ? String(row[headerMap.deliveryNotes] || "").trim() : "";
  const deliveryTerms = headerMap.deliveryTerms !== -1 ? String(row[headerMap.deliveryTerms] || "").trim() : "";
  const email = headerMap.email !== -1 ? String(row[headerMap.email] || "").trim() : "";
  const phone = headerMap.phone !== -1 ? String(row[headerMap.phone] || "").trim() : "";
  const aov = getAovForRow(row, getImportAovMode());
  const assignedRepId = elements.importDefaultRep.value;
  const extraFields = importState.headers.reduce((acc, header, index) => {
    acc[header.trim()] = String(row[index] ?? "").trim();
    return acc;
  }, {});

  return {
    customerId,
    storeName,
    orderTermsLabel,
    mappedOrderTerms,
    fullAddress,
    address1,
    suburb,
    state: stateValue,
    postcode,
    deliveryNotes,
    deliveryTerms,
    email,
    phone,
    averageOrderValue: aov,
    assignedRepId,
    extraFields,
  };
}

async function handleImportRun() {
  if (!canWrite()) {
    showOfflineAlert();
    return;
  }
  const defaultRep = elements.importDefaultRep.value;
  if (!defaultRep) {
    alert("Select a default rep for imported customers.");
    return;
  }
  if (!importState.rows.length) {
    alert("Run preview first to load rows.");
    return;
  }

  const duplicateMode = elements.importDuplicateMode.value;
  let created = 0;
  let updated = 0;
  let skipped = 0;
  let errors = 0;
  const reportRows = [["row_number", "customer_id", "outcome", "message"]];
  const totalRows = importState.rows.length;
  let processed = 0;
  elements.importResults.textContent = "";
  elements.importStatus.textContent = `Starting import for ${totalRows} rows…`;

  const chunkSize = 200;
  for (let start = 0; start < importState.rows.length; start += chunkSize) {
    const chunk = importState.rows.slice(start, start + chunkSize);
    const supabaseBatch = [];
    for (let i = 0; i < chunk.length; i += 1) {
      const rowIndex = start + i;
      const row = chunk[i];
      const customerData = buildCustomerFromRow(row, rowIndex);
      const hasBasicFields = (customerData.customerId || customerData.storeName) && customerData.fullAddress;
      const hasTerms = Boolean(customerData.mappedOrderTerms);
      if (!hasBasicFields || !hasTerms) {
        errors += 1;
        reportRows.push([
          rowIndex + 1,
          customerData.customerId,
          "error",
          "Missing required fields or invalid order terms",
        ]);
        processed += 1;
        continue;
      }

      const existing = getExistingCustomerByExternalId(customerData.customerId);
      if (existing && duplicateMode === "skip") {
        skipped += 1;
        reportRows.push([rowIndex + 1, customerData.customerId, "skipped", "Duplicate customer ID"]);
        processed += 1;
        continue;
      }

      const payload = {
        id: existing ? existing.id : createCustomerId(),
        customerId: customerData.customerId || existing?.customerId || existing?.externalCustomerId || "",
        storeName: customerData.storeName || existing?.storeName || "Unnamed",
        contactName: existing?.contactName || "",
        phone: customerData.phone || existing?.phone || "",
        email: customerData.email || existing?.email || "",
        fullAddress: customerData.fullAddress,
        address1: customerData.address1 || existing?.address1 || existing?.street || "",
        suburb1: customerData.suburb || existing?.suburb1 || existing?.suburb || "",
        state1: customerData.state || existing?.state1 || existing?.state || "",
        postcode1: customerData.postcode || existing?.postcode1 || existing?.postcode || "",
        deliveryNotes: customerData.deliveryNotes || existing?.deliveryNotes || "",
        deliveryTerms: customerData.deliveryTerms || existing?.deliveryTerms || "",
        assignedRepId: customerData.assignedRepId || existing?.assignedRepId || "",
        orderChannel: customerData.mappedOrderTerms || existing?.orderChannel || existing?.channelPreference || "portal",
        orderTermsLabel: orderTermsLabelFromChannel(customerData.mappedOrderTerms) || customerData.orderTermsLabel || existing?.orderTermsLabel || "",
        averageOrderValue:
          customerData.averageOrderValue !== null && customerData.averageOrderValue !== undefined
            ? customerData.averageOrderValue
            : existing?.averageOrderValue ?? null,
        packOffsetDays: existing?.packOffsetDays ?? 0,
        deliveryOffsetDays: existing?.deliveryOffsetDays ?? 1,
        cadenceType: existing?.cadenceType ?? null,
        cadenceDayOfWeek: existing?.cadenceDayOfWeek || "Monday",
        cadenceDaysOfWeek: existing?.cadenceDaysOfWeek || [],
        cadenceEveryNDays: existing?.cadenceEveryNDays || null,
        cadenceNextDueDate: existing?.cadenceNextDueDate || null,
        autoAdvanceCadence: existing?.autoAdvanceCadence || false,
        customerNotes: existing?.customerNotes || "",
        schedule: existing?.schedule || defaultSchedule(),
        extraFields: customerData.extraFields || existing?.extraFields || {},
      };

      try {
        await put("customers", payload);
      } catch (error) {
        console.error("IndexedDB write failed", error);
        setSyncError("Failed to write customer to local storage.");
        return;
      }
      supabaseBatch.push(payload);
      if (existing) {
        updated += 1;
        state.customers = state.customers.map((item) => (item.id === payload.id ? payload : item));
        reportRows.push([rowIndex + 1, customerData.customerId, "updated", "Updated existing customer"]);
      } else {
        created += 1;
        state.customers.push(payload);
        reportRows.push([rowIndex + 1, customerData.customerId, "created", "Created new customer"]);
      }
      processed += 1;
      elements.importStatus.textContent = `Processed ${processed}/${totalRows} rows…`;
    }
    if (supabaseBatch.length) {
      try {
        await syncUpsertBatch("customers", supabaseBatch, mapCustomerToSupabase, {
          onProgress: (done, total) => {
            elements.importStatus.textContent = `Syncing ${processed}/${totalRows} rows • ${done}/${total} customers…`;
          },
        });
        updateConnectionStatus({ status: "Online/Synced", canWrite: true });
        setCloudStatus("synced", toISO());
      } catch (error) {
        console.error("Supabase upsert failed", error);
        handleSupabaseError(error, { context: "Supabase import failed", alertOnOffline: true });
        return;
      }
    }
    await new Promise((resolve) => setTimeout(resolve, 0));
  }

  importState.reportCsv = reportRows.map((row) => row.map(toCsvValue).join(",")).join("\n");
  elements.importResults.textContent = `Import complete. Created ${created}, updated ${updated}, skipped ${skipped}, errors ${errors}.`;
  elements.importStatus.textContent = `Imported ${created + updated} customers and synced to cloud.`;
  renderAll();
}

function downloadImportReport() {
  if (!importState.reportCsv) {
    alert("No report available yet.");
    return;
  }
  const blob = new Blob([importState.reportCsv], { type: "text/csv;charset=utf-8;" });
  const link = document.createElement("a");
  const url = URL.createObjectURL(blob);
  link.setAttribute("href", url);
  link.setAttribute("download", `import-report-${todayKey()}.csv`);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

function openRepModal(rep = {}) {
  showModal(`
    <h2>${rep.id ? "Edit rep" : "Add rep"}</h2>
    <form id="repForm" class="form">
      <label>Name*
        <input name="name" required value="${rep.name || ""}" />
      </label>
      <label>Phone
        <input name="phone" value="${rep.phone || ""}" />
      </label>
      <label>Vehicle
        <input name="vehicle" value="${rep.vehicle || ""}" />
      </label>
      <label><input type="checkbox" name="active" ${rep.active !== false ? "checked" : ""}/> Active</label>
      <div class="form-actions">
        <button class="secondary" type="button" id="cancelRep">Cancel</button>
        <button class="primary" type="submit">Save</button>
      </div>
    </form>
  `);
  disableFormIfOffline("repForm");

  document.getElementById("cancelRep").addEventListener("click", closeModal);
  document.getElementById("repForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    if (!canWrite()) {
      showOfflineAlert();
      return;
    }
    const formData = new FormData(event.target);
    const updated = {
      id: rep.id || uuid(),
      name: formData.get("name").trim(),
      phone: formData.get("phone").trim(),
      vehicle: formData.get("vehicle").trim(),
      active: formData.get("active") === "on",
    };

    if (!updated.name) {
      alert("Name is required.");
      return;
    }

    await put("reps", updated);
    const exists = state.reps.find((item) => item.id === updated.id);
    if (exists) {
      state.reps = state.reps.map((item) => (item.id === updated.id ? updated : item));
    } else {
      state.reps.push(updated);
    }
    closeModal();
    renderAll();
  });
}

function renderCustomerActivity(customerId) {
  if (!customerId) {
    return "<p class=\"muted\">Save the customer to see activity here.</p>";
  }
  const events = state.scheduleEvents
    .filter((event) => event.customerId === customerId)
    .sort((a, b) => (b.date || "").localeCompare(a.date || ""))
    .slice(0, 10);
  if (!events.length) {
    return "<p class=\"muted\">No recent activity yet.</p>";
  }
  return `
    <ul class="muted activity-list">
      ${events
        .map((event) => {
          const kindLabel =
            event.kind === "custom_oneoff"
              ? "One-off"
              : event.kind === "delivery"
                ? "Delivery"
                : event.kind === "pack"
                  ? "Pack"
                  : "Expected order";
          const reasonText = event.skippedReasonText || event.skippedReason || "";
          return `
            <li>
              <strong>${formatDate(event.date)}</strong> • ${kindLabel} • ${event.status}
              ${reasonText ? `• ${reasonText}` : ""}
            </li>
          `;
        })
        .join("")}
    </ul>
  `;
}

function openCustomerModal(customer = {}) {
  const orderChannel = customer.orderChannel || customer.channelPreference || "portal";
  const schedule = customer.schedule || defaultSchedule();
  showModal(`
    <h2>${customer.id ? "Edit customer" : "Add customer"}</h2>
    <form id="customerForm" class="form">
      <label>Customer ID
        <input name="customerId" value="${customer.customerId || customer.externalCustomerId || ""}" />
      </label>
      <label>Store name*
        <input name="storeName" required value="${customer.storeName || ""}" />
      </label>
      <label>Contact name
        <input name="contactName" value="${customer.contactName || ""}" />
      </label>
      <div class="inline-grid">
        <label>Phone
          <input name="phone" value="${customer.phone || ""}" />
        </label>
        <label>Email
          <input name="email" value="${customer.email || ""}" />
        </label>
      </div>
      <label>Full address*
        <textarea name="fullAddress" required>${customer.fullAddress || ""}</textarea>
      </label>
      <div class="inline-grid">
        <label>Address line 1
          <input name="address1" value="${customer.address1 || customer.street || ""}" />
        </label>
        <label>Suburb
          <input name="suburb1" value="${customer.suburb1 || customer.suburb || ""}" />
        </label>
        <label>State
          <input name="state1" value="${customer.state1 || customer.state || ""}" />
        </label>
        <label>Postcode
          <input name="postcode1" value="${customer.postcode1 || customer.postcode || ""}" />
        </label>
      </div>
      <label>Delivery notes
        <textarea name="deliveryNotes">${customer.deliveryNotes || ""}</textarea>
      </label>
      <label>Delivery terms
        <input name="deliveryTerms" value="${customer.deliveryTerms || ""}" />
      </label>
      <label>Assigned rep*
        <select name="assignedRepId" required>
          ${state.reps
            .map(
              (rep) =>
                `<option value="${rep.id}" ${rep.id === customer.assignedRepId ? "selected" : ""}>${rep.name}</option>`
            )
            .join("")}
        </select>
      </label>
      <label>Order channel preference
        <select name="orderChannel">
          <option value="portal" ${orderChannel === "portal" ? "selected" : ""}>Order through portal</option>
          <option value="sms" ${orderChannel === "sms" ? "selected" : ""}>SMS order</option>
          <option value="email" ${orderChannel === "email" ? "selected" : ""}>Email order</option>
          <option value="phone" ${orderChannel === "phone" ? "selected" : ""}>Phone call order</option>
          <option value="rep_in_person" ${orderChannel === "rep_in_person" ? "selected" : ""}>Rep in-person order</option>
        </select>
      </label>
      <label>Average order value (AUD)
        <input name="averageOrderValue" type="number" value="${customer.averageOrderValue ?? ""}" />
      </label>
      <label>Customer notes
        <textarea name="customerNotes" rows="3">${customer.customerNotes || ""}</textarea>
      </label>
      <div class="card">
        <h3>Scheduling</h3>
        <label>How do we get their order? ${helpIcon("Choose how the order arrives. Example: We get the order = rep calls/visits; Customer puts order = they submit on their own. Usual plan only; you can still add one-off changes.")} 
          <select name="scheduleMode">
            <option value="">Select...</option>
            <option value="WE_GET_ORDER" ${schedule.mode === "WE_GET_ORDER" ? "selected" : ""}>We get the order (sales rep takes it)</option>
            <option value="THEY_PUT_ORDER" ${schedule.mode === "THEY_PUT_ORDER" ? "selected" : ""}>Customer puts their own order in</option>
          </select>
        </label>
        <label>How often? ${helpIcon("Weekly happens every week. Fortnightly/Every 3 weeks uses the anchor date to align weeks. Usual plan only; you can still add one-off changes.")} 
          <select name="scheduleFrequency">
            <option value="">Select...</option>
            <option value="WEEKLY" ${schedule.frequency === "WEEKLY" ? "selected" : ""}>Weekly</option>
            <option value="FORTNIGHTLY" ${schedule.frequency === "FORTNIGHTLY" ? "selected" : ""}>Fortnightly (every 2 weeks)</option>
            <option value="EVERY_3_WEEKS" ${schedule.frequency === "EVERY_3_WEEKS" ? "selected" : ""}>Every 3 weeks</option>
          </select>
        </label>
        <label id="scheduleAnchorWrap">Anchor date ${helpIcon("Pick a date that represents the last cycle week for this customer. The app uses it to align fortnightly/3-week schedules. Example: if their last order week started on 3 Jan, pick 2025-01-03.")} 
          <input name="scheduleAnchorDate" type="date" value="${schedule.anchorDate || ""}" />
        </label>
        <div id="scheduleWeGet">
          <label>Order day ${helpIcon("Day we usually take their order (call/visit/text). Example: If order day is Mon, pack Mon, deliver Tue. Usual plan only; you can still add one-off changes.")} 
            <select name="orderDay1">
              <option value="">Select day</option>
              ${dayNames
                .map((day, index) => `<option value="${index}" ${Number(schedule.orderDay1) === index ? "selected" : ""}>${day}</option>`)
                .join("")}
            </select>
          </label>
          <label>Pack day ${helpIcon("Day we usually pack for this order. If blank, it defaults to the order day. Usual plan only; you can still add one-off changes.")} 
            <select name="packDay1">
              <option value="">Default same as order day</option>
              ${dayNames
                .map((day, index) => `<option value="${index}" ${Number(schedule.packDay1) === index ? "selected" : ""}>${day}</option>`)
                .join("")}
            </select>
          </label>
          <label>Delivery day ${helpIcon("Day we deliver. If blank, default is the next day after order day. Usual plan only; you can still add one-off changes.")} 
            <select name="deliverDay1">
              <option value="">Default next day</option>
              ${dayNames
                .map((day, index) => `<option value="${index}" ${Number(schedule.deliverDay1) === index ? "selected" : ""}>${day}</option>`)
                .join("")}
            </select>
          </label>
          <label><input type="checkbox" name="isBiWeeklySecondRun" ${schedule.isBiWeeklySecondRun ? "checked" : ""}/> Biweekly customer (2 runs per week) ${helpIcon("If they order/deliver twice per week, enable this to set a second order + pack + delivery day. Usual plan only; you can still add one-off changes.")}</label>
          <div id="scheduleSecondRun">
            <label>Second order day ${helpIcon("Second day we take their order. Example: Tue + Fri. Usual plan only; you can still add one-off changes.")} 
              <select name="orderDay2">
                <option value="">Select day</option>
                ${dayNames
                  .map((day, index) => `<option value="${index}" ${Number(schedule.orderDay2) === index ? "selected" : ""}>${day}</option>`)
                  .join("")}
              </select>
            </label>
            <label>Second pack day ${helpIcon("Second pack day. If blank, it defaults to the second order day. Usual plan only; you can still add one-off changes.")} 
              <select name="packDay2">
                <option value="">Default same as order day</option>
                ${dayNames
                  .map((day, index) => `<option value="${index}" ${Number(schedule.packDay2) === index ? "selected" : ""}>${day}</option>`)
                  .join("")}
              </select>
            </label>
            <label>Second delivery day ${helpIcon("Second delivery day. If blank, default is next day after second order day. Usual plan only; you can still add one-off changes.")} 
              <select name="deliverDay2">
                <option value="">Default next day</option>
                ${dayNames
                  .map((day, index) => `<option value="${index}" ${Number(schedule.deliverDay2) === index ? "selected" : ""}>${day}</option>`)
                  .join("")}
              </select>
            </label>
          </div>
        </div>
        <div id="scheduleTheyPut">
          <label>Which day(s) do they usually place orders? ${helpIcon("Select one or multiple days. Example: Mon + Thu. Usual plan only; you can still add one-off changes.")} 
            ${buildDayButtons(schedule.customerOrderDays || [], "customerOrderDays")}
          </label>
          <label>Pack day(s) ${helpIcon("Select pack days. If none selected, pack defaults to the order day. Usual plan only; you can still add one-off changes.")} 
            ${buildDayButtons(schedule.packDays || [], "packDays")}
          </label>
          <label>Delivery day(s) ${helpIcon("Select delivery days. If none selected, delivery defaults to the day after order. Usual plan only; you can still add one-off changes.")} 
            ${buildDayButtons(schedule.deliverDays || [], "deliverDays")}
          </label>
        </div>
      </div>
      <div class="card">
        <h3>Recent activity</h3>
        ${renderCustomerActivity(customer.id)}
      </div>
      <div class="form-actions">
        <button class="secondary" type="button" id="cancelCustomer">Cancel</button>
        <button class="primary" type="submit">Save</button>
      </div>
    </form>
  `);
  disableFormIfOffline("customerForm");

  document.getElementById("cancelCustomer").addEventListener("click", closeModal);
  const scheduleModeSelect = document.querySelector("select[name='scheduleMode']");
  const scheduleFrequencySelect = document.querySelector("select[name='scheduleFrequency']");
  const anchorWrap = document.getElementById("scheduleAnchorWrap");
  const weGetSection = document.getElementById("scheduleWeGet");
  const theyPutSection = document.getElementById("scheduleTheyPut");
  const secondRunWrap = document.getElementById("scheduleSecondRun");
  const secondRunToggle = document.querySelector("input[name='isBiWeeklySecondRun']");

  const updateScheduleVisibility = () => {
    const mode = scheduleModeSelect.value;
    const frequency = scheduleFrequencySelect.value;
    const showAnchor = frequency === "FORTNIGHTLY" || frequency === "EVERY_3_WEEKS";
    anchorWrap.style.display = showAnchor ? "grid" : "none";
    weGetSection.style.display = mode === "WE_GET_ORDER" ? "grid" : "none";
    theyPutSection.style.display = mode === "THEY_PUT_ORDER" ? "grid" : "none";
    secondRunWrap.style.display = secondRunToggle.checked ? "grid" : "none";
  };

  const toggleDayButton = (button) => {
    button.classList.toggle("selected");
  };

  document.querySelectorAll(".day-button").forEach((button) => {
    button.addEventListener("click", () => toggleDayButton(button));
  });

  const getSelectedDays = (role) => {
    const container = document.querySelector(`[data-day-role="${role}"]`);
    if (!container) return [];
    return Array.from(container.querySelectorAll(".day-button.selected")).map((btn) => Number(btn.dataset.day));
  };

  scheduleModeSelect.addEventListener("change", updateScheduleVisibility);
  scheduleFrequencySelect.addEventListener("change", updateScheduleVisibility);
  secondRunToggle.addEventListener("change", updateScheduleVisibility);
  updateScheduleVisibility();

  document.getElementById("customerForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    if (!canWrite()) {
      showOfflineAlert();
      return;
    }
    const formData = new FormData(event.target);
    const address1 = formData.get("address1").trim();
    const suburb1 = formData.get("suburb1").trim();
    const state1 = formData.get("state1").trim();
    const postcode1 = formData.get("postcode1").trim();
    const fullAddressValue = formData.get("fullAddress").trim() || [address1, suburb1, state1, postcode1].filter(Boolean).join(", ");
    const orderChannel = formData.get("orderChannel") || "portal";
    const scheduleMode = formData.get("scheduleMode") || null;
    const scheduleFrequency = formData.get("scheduleFrequency") || null;
    const scheduleAnchorDate = formData.get("scheduleAnchorDate") || null;
    if ((scheduleFrequency === "FORTNIGHTLY" || scheduleFrequency === "EVERY_3_WEEKS") && !scheduleAnchorDate) {
      alert("Anchor date is required for fortnightly or every 3 weeks schedules.");
      return;
    }
    const updated = {
      id: customer.id || createCustomerId(),
      customerId: formData.get("customerId").trim(),
      storeName: formData.get("storeName").trim(),
      contactName: formData.get("contactName").trim(),
      phone: formData.get("phone").trim(),
      email: formData.get("email").trim(),
      fullAddress: fullAddressValue,
      address1,
      suburb1,
      state1,
      postcode1,
      deliveryNotes: formData.get("deliveryNotes").trim(),
      deliveryTerms: formData.get("deliveryTerms").trim(),
      assignedRepId: formData.get("assignedRepId"),
      orderChannel,
      orderTermsLabel: orderTermsLabelFromChannel(orderChannel),
      averageOrderValue: formData.get("averageOrderValue")
        ? Number(formData.get("averageOrderValue"))
        : null,
      packOffsetDays: customer.packOffsetDays ?? 0,
      deliveryOffsetDays: customer.deliveryOffsetDays ?? 1,
      cadenceType: customer.cadenceType ?? null,
      cadenceDayOfWeek: customer.cadenceDayOfWeek,
      cadenceDaysOfWeek: customer.cadenceDaysOfWeek || [],
      cadenceEveryNDays: customer.cadenceEveryNDays || null,
      cadenceNextDueDate: customer.cadenceNextDueDate || null,
      autoAdvanceCadence: customer.autoAdvanceCadence || false,
      customerNotes: formData.get("customerNotes").trim(),
      schedule: {
        mode: scheduleMode,
        frequency: scheduleFrequency,
        orderDay1: formData.get("orderDay1") ? Number(formData.get("orderDay1")) : null,
        deliverDay1: formData.get("deliverDay1") ? Number(formData.get("deliverDay1")) : null,
        packDay1: formData.get("packDay1") ? Number(formData.get("packDay1")) : null,
        isBiWeeklySecondRun: formData.get("isBiWeeklySecondRun") === "on",
        orderDay2: formData.get("orderDay2") ? Number(formData.get("orderDay2")) : null,
        deliverDay2: formData.get("deliverDay2") ? Number(formData.get("deliverDay2")) : null,
        packDay2: formData.get("packDay2") ? Number(formData.get("packDay2")) : null,
        customerOrderDays: getSelectedDays("customerOrderDays"),
        deliverDays: getSelectedDays("deliverDays"),
        packDays: getSelectedDays("packDays"),
        anchorDate: scheduleAnchorDate,
      },
      extraFields: customer.extraFields || {},
    };

    if (!updated.storeName || !updated.fullAddress || !updated.assignedRepId) {
      alert("Store name, address, and assigned rep are required.");
      return;
    }

    try {
      await syncUpsertCustomer(updated);
      await put("customers", updated);
      updateConnectionStatus({ status: "Online/Synced", canWrite: true });
      setCloudStatus("synced", toISO());
    } catch (error) {
      console.error("Supabase upsert failed", error);
      handleSupabaseError(error, { context: "Supabase upsert failed", alertOnOffline: true });
      return;
    }
    const exists = state.customers.find((item) => item.id === updated.id);
    if (exists) {
      state.customers = state.customers.map((item) => (item.id === updated.id ? updated : item));
    } else {
      state.customers.push(updated);
    }
    closeModal();
    renderAll();
  });
}

function orderLinesForm(orderLines = []) {
  return `
    <div id="orderLines">
      ${orderLines
        .map(
          (line, index) => `
        <div class="inline-grid order-line" data-index="${index}">
          <input name="sku_${index}" placeholder="Item" value="${line.skuOrItemName || ""}" />
          <input name="qty_${index}" placeholder="Qty" type="number" value="${line.qty || 1}" />
          <input name="note_${index}" placeholder="Notes" value="${line.notes || ""}" />
        </div>
      `
        )
        .join("")}
    </div>
    <button id="addLineBtn" class="secondary" type="button">Add line</button>
  `;
}

function openOrderModal(order = {}) {
  const customer = customerById(order.customerId) || {};
  const defaultChannel = order.channel || customer.orderChannel || customer.channelPreference || "portal";
  const defaultTaskType = order.taskType || "delivery";
  showModal(`
    <h2>${order.id ? "Edit order" : "Quick add order"}</h2>
    <form id="orderForm" class="form">
      <label>Customer*
        <select name="customerId" required>
          <option value="">Select customer</option>
          ${state.customers
            .map(
              (cust) =>
                `<option value="${cust.id}" ${cust.id === order.customerId ? "selected" : ""}>${cust.storeName}</option>`
            )
            .join("")}
        </select>
        <span id="orderCustomerHint" class="muted"></span>
      </label>
      <label>Channel*
        <select name="channel" required>
          <option value="portal" ${defaultChannel === "portal" ? "selected" : ""}>Portal</option>
          <option value="sms" ${defaultChannel === "sms" ? "selected" : ""}>SMS</option>
          <option value="email" ${defaultChannel === "email" ? "selected" : ""}>Email</option>
          <option value="phone" ${defaultChannel === "phone" ? "selected" : ""}>Phone</option>
          <option value="rep_in_person" ${defaultChannel === "rep_in_person" ? "selected" : ""}>Sales rep</option>
        </select>
      </label>
      <div class="inline-grid">
        <label>Received at*
          <input name="receivedAt" type="datetime-local" value="${order.receivedAt ? order.receivedAt.slice(0, 16) : toISO().slice(0, 16)}" />
        </label>
        <label>Received by
          <input name="receivedBy" value="${order.receivedBy || ""}" />
        </label>
      </div>
      <div class="inline-grid">
        <label>Pack offset (days)
          <input name="packOffsetDays" type="number" value="${order.packOffsetDays ?? customer.packOffsetDays ?? 0}" />
        </label>
        <label>Delivery offset (days)
          <input name="deliveryOffsetDays" type="number" value="${order.deliveryOffsetDays ?? customer.deliveryOffsetDays ?? 1}" />
        </label>
      </div>
      <label>Assigned rep
        <select name="assignedRepId">
          ${state.reps
            .map(
              (rep) =>
                `<option value="${rep.id}" ${rep.id === (order.assignedRepId || customer.assignedRepId) ? "selected" : ""}>${rep.name}</option>`
            )
            .join("")}
        </select>
      </label>
      <label>Status
        <select name="status">
          ${Object.keys(statusLabels)
            .map(
              (status) =>
                `<option value="${status}" ${status === (order.status || "received") ? "selected" : ""}>${statusLabels[status]}</option>`
            )
            .join("")}
        </select>
      </label>
      <label>Task type
        <select name="taskType">
          <option value="delivery" ${defaultTaskType === "delivery" ? "selected" : ""}>Delivery</option>
          <option value="pickup" ${defaultTaskType === "pickup" ? "selected" : ""}>Pickup</option>
          <option value="packing" ${defaultTaskType === "packing" ? "selected" : ""}>Packing</option>
          <option value="other" ${defaultTaskType === "other" ? "selected" : ""}>Other</option>
        </select>
      </label>
      <label>Order lines
        ${orderLinesForm(order.orderLines || [])}
      </label>
      <label>Internal notes
        <textarea name="internalNotes">${order.internalNotes || ""}</textarea>
      </label>
      <label>Service time (minutes)
        <input name="serviceTime" type="number" value="${order.serviceTime || ""}" />
      </label>
      <div class="form-actions">
        ${order.id ? '<button class="secondary" id="recalcTasksBtn" type="button">Recalculate tasks</button>' : ""}
        <button class="secondary" type="button" id="cancelOrder">Cancel</button>
        <button class="primary" type="submit">Save</button>
      </div>
    </form>
  `);
  disableFormIfOffline("orderForm");

  document.getElementById("cancelOrder").addEventListener("click", closeModal);

  const addLineBtn = document.getElementById("addLineBtn");
  addLineBtn?.addEventListener("click", () => {
    const container = document.getElementById("orderLines");
    const index = container.children.length;
    const line = document.createElement("div");
    line.className = "inline-grid order-line";
    line.dataset.index = index;
    line.innerHTML = `
      <input name="sku_${index}" placeholder="Item" />
      <input name="qty_${index}" placeholder="Qty" type="number" value="1" />
      <input name="note_${index}" placeholder="Notes" />
    `;
    container.appendChild(line);
  });

  const customerSelect = document.querySelector("select[name='customerId']");
  const hint = document.getElementById("orderCustomerHint");
  const updateHint = () => {
    const selected = customerById(customerSelect.value);
    if (selected && selected.averageOrderValue !== null && selected.averageOrderValue !== undefined) {
      hint.textContent = `Avg order value: $${selected.averageOrderValue}`;
    } else {
      hint.textContent = "";
    }
  };
  customerSelect.addEventListener("change", updateHint);
  updateHint();

  const recalcBtn = document.getElementById("recalcTasksBtn");
  recalcBtn?.addEventListener("click", async () => {
    if (!canWrite()) {
      showOfflineAlert();
      return;
    }
    const formData = new FormData(document.getElementById("orderForm"));
    const updatedOrder = buildOrderFromForm(Object.fromEntries(formData));
    updatedOrder.id = order.id;
    await updateOrderAndTasks(updatedOrder, true);
    closeModal();
    renderAll();
  });

  document.getElementById("orderForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    if (!canWrite()) {
      showOfflineAlert();
      return;
    }
    const formData = new FormData(event.target);
    const data = Object.fromEntries(formData);
    const updatedOrder = buildOrderFromForm(data, order.id);

    if (!updatedOrder.customerId || !updatedOrder.channel || !updatedOrder.receivedAt) {
      alert("Customer, channel, and received date are required.");
      return;
    }

    if (order.id) {
      await updateOrderAndTasks(updatedOrder, false);
    } else {
      await saveOrder(updatedOrder, true);
    }
    closeModal();
    renderAll();
  });
}

function buildOrderFromForm(data, orderId = null) {
  const customer = customerById(data.customerId) || {};
  const packOffsetDays = Number(data.packOffsetDays || customer.packOffsetDays || 0);
  const deliveryOffsetDays = Number(data.deliveryOffsetDays || customer.deliveryOffsetDays || 1);
  const dueDates = computeDueDates(data.receivedAt, packOffsetDays, deliveryOffsetDays);
  const orderLines = [];
  const taskType = data.taskType || "delivery";

  Object.keys(data)
    .filter((key) => key.startsWith("sku_"))
    .forEach((key) => {
      const index = key.split("_")[1];
      const sku = data[`sku_${index}`];
      const qty = Number(data[`qty_${index}`] || 1);
      const notes = data[`note_${index}`];
      if (sku) {
        orderLines.push({ skuOrItemName: sku, qty, notes });
      }
    });

  return {
    id: orderId || uuid(),
    customerId: data.customerId,
    channel: data.channel,
    receivedAt: new Date(data.receivedAt).toISOString(),
    receivedBy: data.receivedBy || "",
    orderLines,
    internalNotes: data.internalNotes || "",
    status: data.status || "received",
    assignedRepId: data.assignedRepId || customer.assignedRepId,
    packOffsetDays,
    deliveryOffsetDays,
    packDueDate: dueDates.packDueDate,
    deliveryDueDate: dueDates.deliveryDueDate,
    serviceTime: data.serviceTime ? Number(data.serviceTime) : null,
    taskType,
  };
}

async function saveOrder(order, createTasks) {
  await put("orders", order);
  state.orders.push(order);
  if (createTasks) {
    await createOrderTasks(order);
  }
  const customer = customerById(order.customerId);
  if (customer?.autoAdvanceCadence && customer.cadenceType === "fortnightly") {
    customer.cadenceNextDueDate = addDays(order.deliveryDueDate, 14);
    await put("customers", customer);
  }
}

async function updateOrderAndTasks(order, recalcTasks) {
  await put("orders", order);
  state.orders = state.orders.map((item) => (item.id === order.id ? order : item));
  if (recalcTasks) {
    await removeTasksForOrder(order.id);
    await createOrderTasks(order);
  }
  if (order.status === "packed") {
    const packTask = state.tasks.find((task) => task.orderId === order.id && task.type === "pack");
    if (packTask) {
      packTask.status = "done";
      await put("tasks", packTask);
    }
  }
  if (order.status === "delivered") {
    const deliveryTask = state.tasks.find((task) => task.orderId === order.id && task.type === "deliver");
    if (deliveryTask) {
      deliveryTask.status = "done";
      await put("tasks", deliveryTask);
    }
  }
}

function renderBackupStatus() {
  if (!state.settings.app.lastBackupAt) {
    elements.backupStatus.textContent = "No backup yet.";
    return;
  }
  elements.backupStatus.textContent = `Last backup: ${formatDateTime(state.settings.app.lastBackupAt)}`;
}

function openHelpModal() {
  showModal(`
    <h2>User guide</h2>
    <p>This planner keeps Orchard Valley orders flowing from order → pack → deliver.</p>
    <ul>
      <li><strong>Start with reps and customers.</strong> Assign each customer a rep and cadence.</li>
      <li><strong>Quick add orders</strong> from any channel. Tasks are created automatically.</li>
      <li><strong>Today view</strong> shows pack today + deliver today/tomorrow. Mark done with one click.</li>
      <li><strong>Expected orders</strong> uses cadence settings for planning.</li>
      <li><strong>Export deliveries</strong> by date/rep with configurable column mapping.</li>
      <li><strong>Backup & restore</strong> for safe offline storage.</li>
    </ul>
    <p class="muted">Timezone: Australia/Melbourne. Data is stored locally in your browser.</p>
  `);
}

async function handleBackup() {
  const data = {
    reps: state.reps,
    customers: state.customers,
    orders: state.orders,
    tasks: state.tasks,
    scheduleEvents: state.scheduleEvents,
    oneOffItems: state.oneOffItems,
    settings: state.settings,
  };
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `orchard-valley-backup-${todayKey()}.json`;
  link.click();
  URL.revokeObjectURL(url);
  state.settings.app.lastBackupAt = toISO();
  await saveSettings();
  renderBackupStatus();
}

async function handleRestore() {
  if (!canWrite()) {
    showOfflineAlert();
    return;
  }
  const file = elements.restoreInput.files[0];
  if (!file) return;
  if (!confirm("Restore backup? This will overwrite current data.")) return;
  const text = await file.text();
  const data = JSON.parse(text);
  const importedCustomers = (data.customers || []).map((customer) => ({
    ...customer,
    id: customer.id || createCustomerId(),
  }));
  for (const store of ["reps", "customers", "orders", "tasks", "schedule_events", "one_off_items", "settings"]) {
    await clearStore(store);
  }
  await bulkPut("reps", data.reps || []);
  await bulkPut("customers", importedCustomers);
  await bulkPut("orders", data.orders || []);
  await bulkPut("tasks", data.tasks || []);
  await bulkPut("schedule_events", data.scheduleEvents || []);
  await bulkPut("one_off_items", data.oneOffItems || []);
  if (data.settings?.app) {
    await put("settings", data.settings.app);
  }
  await loadState();
  const customers = importedCustomers;
  const events = data.scheduleEvents || [];
  elements.backupStatus.textContent = `Restoring ${customers.length} customers…`;
  try {
    await syncUpsertBatch("customers", customers, mapCustomerToSupabase, {
      onProgress: (done, total) => {
        elements.backupStatus.textContent = `Syncing customers ${done}/${total}…`;
      },
    });
    await syncUpsertBatch("schedule_events", events, mapScheduleEventToSupabase, {
      onProgress: (done, total) => {
        elements.backupStatus.textContent = `Syncing schedule events ${done}/${total}…`;
      },
    });
    updateConnectionStatus({ status: "Online/Synced", canWrite: true });
    setCloudStatus("synced", toISO());
    elements.backupStatus.textContent = `Imported ${customers.length} customers and synced to cloud.`;
  } catch (error) {
    console.error("Supabase sync failed", error);
    handleSupabaseError(error, { context: "Supabase restore failed", alertOnOffline: true });
    elements.backupStatus.textContent = `Restore complete locally. Supabase sync failed.`;
  }
  renderAll();
}

function exportRecordsForRange() {
  const start = elements.exportStart.value || todayKey();
  let end = elements.exportEnd.value || start;
  if (elements.exportIncludeTomorrow.checked) {
    end = addDays(end, 1);
  }
  const repFilter = elements.exportRep.value;
  const includeCompleted = elements.exportIncludeCompleted.checked;
  return buildSpokenitExportRecords({
    tasks: state.tasks,
    orders: state.orders,
    customers: state.customers,
    reps: state.reps,
    start,
    end,
    repFilter,
    includeCompleted,
  });
}

async function handleExport() {
  const records = exportRecordsForRange();
  const preset = getActivePreset();
  if (!preset) return;
  const csv = buildCsv(records, preset.columns);
  downloadCsv(`deliveries-${elements.exportStart.value || todayKey()}.csv`, csv);
  elements.exportSummary.textContent = `${records.length} stops exported.`;
}

function setupExportEvents() {
  elements.mappingPresetSelect.addEventListener("change", () => {
    state.settings.app.activeExportPresetId = elements.mappingPresetSelect.value;
    renderMappingEditor();
  });

  elements.newPresetBtn.addEventListener("click", () => {
    const name = prompt("Preset name?");
    if (!name) return;
    const id = uuid();
    state.settings.app.exportPresets.push({ id, name, columns: [...defaultColumns] });
    state.settings.app.activeExportPresetId = id;
    renderMappingPresets();
  });

  elements.savePresetBtn.addEventListener("click", async () => {
    await saveSettings();
    alert("Preset saved.");
  });

  elements.exportCsvBtn.addEventListener("click", handleExport);
}

async function loadSampleData() {
  if (!canWrite()) {
    showOfflineAlert();
    return;
  }
  if (!confirm("Load sample data? This will add demo reps, customers, and orders.")) return;

  const repA = { id: uuid(), name: "Jamie Cole", phone: "0412 555 012", vehicle: "Van 1", active: true };
  const repB = { id: uuid(), name: "Riley Singh", phone: "0412 555 013", vehicle: "Van 2", active: true };
  const customers = [
    {
      id: uuid(),
      customerId: "CUST-1001",
      storeName: "Yarra Fresh Foods",
      contactName: "Mia",
      phone: "03 9123 4567",
      email: "hello@yarrafresh.com",
      fullAddress: "12 River Rd, Richmond, VIC 3121",
      address1: "12 River Rd",
      suburb1: "Richmond",
      state1: "VIC",
      postcode1: "3121",
      deliveryNotes: "Deliver to back dock",
      deliveryTerms: "Leave at loading dock",
      assignedRepId: repA.id,
      orderChannel: "portal",
      orderTermsLabel: "Order through portal",
      averageOrderValue: 420,
      cadenceType: null,
      cadenceDayOfWeek: "Monday",
      packOffsetDays: 0,
      deliveryOffsetDays: 1,
      autoAdvanceCadence: true,
      customerNotes: "Prefers morning call-ins.",
      schedule: {
        mode: "WE_GET_ORDER",
        frequency: "WEEKLY",
        orderDay1: 1,
        deliverDay1: 2,
        packDay1: 1,
        isBiWeeklySecondRun: false,
        orderDay2: null,
        deliverDay2: null,
        packDay2: null,
        customerOrderDays: [],
        deliverDays: [],
        packDays: [],
        anchorDate: null,
      },
      extraFields: {},
    },
    {
      id: uuid(),
      customerId: "CUST-1002",
      storeName: "Mornington Market",
      contactName: "Luke",
      phone: "03 9012 3400",
      email: "orders@morningtonmarket.com",
      fullAddress: "88 Coast Hwy, Mornington, VIC 3931",
      address1: "88 Coast Hwy",
      suburb1: "Mornington",
      state1: "VIC",
      postcode1: "3931",
      deliveryNotes: "Call on arrival",
      deliveryTerms: "Signature required",
      assignedRepId: repB.id,
      orderChannel: "phone",
      orderTermsLabel: "Phone call order",
      averageOrderValue: 380,
      cadenceType: null,
      cadenceDaysOfWeek: ["Tuesday", "Friday"],
      packOffsetDays: 0,
      deliveryOffsetDays: 1,
      autoAdvanceCadence: false,
      customerNotes: "Call ahead for gate code.",
      schedule: {
        mode: "THEY_PUT_ORDER",
        frequency: "FORTNIGHTLY",
        orderDay1: null,
        deliverDay1: null,
        packDay1: null,
        isBiWeeklySecondRun: false,
        orderDay2: null,
        deliverDay2: null,
        packDay2: null,
        customerOrderDays: [2],
        deliverDays: [],
        packDays: [],
        anchorDate: todayKey(),
      },
      extraFields: {},
    },
  ];

  await bulkPut("reps", [repA, repB]);
  await bulkPut("customers", customers);
    try {
      for (const customer of customers) {
        await syncUpsertCustomer(customer);
      }
      updateConnectionStatus({ status: "Online/Synced", canWrite: true });
      setCloudStatus("synced", toISO());
    } catch (error) {
    console.error("Supabase upsert failed", error);
    handleSupabaseError(error, { context: "Supabase upsert failed", alertOnOffline: true });
    return;
  }

  const order = buildOrderFromForm({
    customerId: customers[0].id,
    channel: "portal",
    receivedAt: toISO(),
    packOffsetDays: 0,
    deliveryOffsetDays: 1,
    assignedRepId: repA.id,
    status: "received",
    sku_0: "Mixed apples",
    qty_0: 10,
    note_0: "",
    internalNotes: "Urgent",
  });

  await saveOrder(order, true);
  await loadState();
  renderAll();
}

function renderAll() {
  renderRepFilters();
  syncScheduleViewControls();
  renderDashboard();
  renderOrders();
  renderCustomers();
  renderReps();
  renderExpected();
  renderSchedule();
  renderBackupStatus();
}

function clearAppState() {
  state.reps = [];
  state.customers = [];
  state.orders = [];
  state.tasks = [];
  state.scheduleEvents = [];
  state.oneOffItems = [];
}

async function handleSession(session) {
  state.session = session;
  if (!session) {
    clearAppState();
    showLoginScreen();
    return;
  }
  showAppShell();
  updateConnectionStatus({
    online: navigator.onLine,
    canWrite: navigator.onLine,
    email: session.user?.email || "",
    status: navigator.onLine ? "Online/Synced" : "Offline/Local",
  });
  const loadResult = await loadFromSupabase();
  if (loadResult?.status === "error") {
    await loadState({ useLocalCustomers: false, useLocalScheduleEvents: false });
  } else {
    await loadState();
  }
  renderAll();
}

function setupAuthEvents() {
  if (!supabaseAvailable) {
    setLoginError("Cloud sign-in is unavailable. Continue in local-only mode.");
    if (elements.loginForm) {
      elements.loginForm
        .querySelectorAll("input, button")
        .forEach((input) => (input.disabled = true));
    }
    return;
  }
  if (elements.loginForm) {
    elements.loginForm.addEventListener("submit", async (event) => {
      event.preventDefault();
      setLoginError("");
      setLoginLoading(true);
      const email = elements.loginEmail.value.trim();
      const password = elements.loginPassword.value;
      const { error } = await signIn(email, password);
      if (error) {
        setLoginError(error.message || "Unable to sign in.");
      }
      setLoginLoading(false);
    });
  }
  if (elements.loginShowPassword && elements.loginPassword) {
    elements.loginShowPassword.addEventListener("change", () => {
      elements.loginPassword.type = elements.loginShowPassword.checked ? "text" : "password";
    });
  }
  if (elements.logoutBtn) {
    elements.logoutBtn.addEventListener("click", async () => {
      await signOut();
    });
  }
  supabase.auth.onAuthStateChange(async (_event, session) => {
    await handleSession(session);
  });
  window.addEventListener("online", async () => {
    updateConnectionStatus({
      online: true,
      canWrite: true,
      status: "Online/Synced",
    });
    if (state.session) {
      const loadResult = await loadFromSupabase();
      if (loadResult?.status === "error") {
        await loadState({ useLocalCustomers: false, useLocalScheduleEvents: false });
      } else {
        await loadState();
      }
      renderAll();
    }
  });
  window.addEventListener("offline", () => {
    updateConnectionStatus({
      online: false,
      canWrite: false,
      status: "Offline/Local",
    });
  });
}

function setupEvents() {
  on(elements.dashboardRepFilter, "change", renderDashboard);
  on(elements.dashboardSearch, "input", renderDashboard);
  on(elements.customersSearch, "input", renderCustomers);
  on(elements.customersRepFilter, "change", renderCustomers);
  on(elements.customersModeFilter, "change", renderCustomers);
  on(elements.customersFrequencyFilter, "change", renderCustomers);
  on(elements.customersChannelFilter, "change", renderCustomers);
  on(elements.customersScheduleFilter, "change", renderCustomers);
  on(elements.addOneOffTodayBtn, "click", openOneOffModal);
  on(elements.addOneOffScheduleBtn, "click", openOneOffModal);
  on(elements.newOrderBtn, "click", () => openOrderModal());
  on(elements.newCustomerBtn, "click", () => {
    if (!state.reps.length) {
      alert("Add a rep before creating customers.");
      return;
    }
    openCustomerModal();
  });
  on(elements.newRepBtn, "click", () => openRepModal());
  on(elements.sampleDataBtn, "click", loadSampleData);
  on(elements.helpBtn, "click", openHelpModal);
  on(elements.accountLogoutBtn, "click", async () => {
    await signOut();
  });
  on(elements.modalClose, "click", closeModal);
  on(elements.snackbarUndo, "click", async () => {
    if (lastUndo) {
      await lastUndo();
    }
    hideSnackbar();
  });
  on(elements.backupBtn, "click", handleBackup);
  on(elements.restoreBtn, "click", handleRestore);
  on(elements.importPreviewBtn, "click", handleImportPreview);
  on(elements.importRunBtn, "click", handleImportRun);
  on(elements.importReportBtn, "click", downloadImportReport);
  on(elements.importDefaultRep, "change", renderImportPreview);
  on(elements.importAovColumn, "change", renderImportPreview);
  on(elements.importAovValue, "input", renderImportPreview);
  onDoc("input[name='importAovMode']", "change", () => renderImportPreview());
  on(elements.scheduleRepFilter, "change", renderSchedule);
  onDoc(".schedule-view-button", "click", async (_event, button) => {
    state.settings.app.scheduleView.viewMode = button.dataset.view;
    if (button.dataset.view === "week") {
      state.settings.app.scheduleView.activeDayKey =
        state.settings.app.scheduleView.anchorDate || todayKey();
    }
    await saveSettings();
    syncScheduleViewControls();
    renderSchedule();
  });
  on(elements.schedulePrevBtn, "click", async () => {
    const viewMode = state.settings.app.scheduleView.viewMode;
    const current = state.settings.app.scheduleView.anchorDate || todayKey();
    const nextDate =
      viewMode === "month"
        ? addMonthsToDateKey(current, -1)
        : addDays(current, viewMode === "week" ? -7 : -1);
    state.settings.app.scheduleView.anchorDate = nextDate;
    state.settings.app.scheduleView.activeDayKey = nextDate;
    await saveSettings();
    renderSchedule();
  });
  on(elements.scheduleNextBtn, "click", async () => {
    const viewMode = state.settings.app.scheduleView.viewMode;
    const current = state.settings.app.scheduleView.anchorDate || todayKey();
    const nextDate =
      viewMode === "month"
        ? addMonthsToDateKey(current, 1)
        : addDays(current, viewMode === "week" ? 7 : 1);
    state.settings.app.scheduleView.anchorDate = nextDate;
    state.settings.app.scheduleView.activeDayKey = nextDate;
    await saveSettings();
    renderSchedule();
  });
  on(elements.scheduleTodayBtn, "click", async () => {
    const today = todayKey();
    state.settings.app.scheduleView.anchorDate = today;
    state.settings.app.scheduleView.activeDayKey = today;
    await saveSettings();
    renderSchedule();
  });
  on(elements.scheduleDatePicker, "change", async () => {
    const nextDate = elements.scheduleDatePicker.value || todayKey();
    state.settings.app.scheduleView.anchorDate = nextDate;
    state.settings.app.scheduleView.activeDayKey = nextDate;
    await saveSettings();
    renderSchedule();
  });
  on(elements.scheduleSearch, "input", async () => {
    state.settings.app.scheduleView.searchTerm = elements.scheduleSearch.value.trim();
    await saveSettings();
    renderSchedule();
  });
  on(elements.scheduleSelectModeBtn, "click", () => {
    document.body.classList.toggle("select-mode");
    const isSelectMode = document.body.classList.contains("select-mode");
    elements.scheduleSelectModeBtn.textContent = isSelectMode ? "Done" : "Select";
    if (!isSelectMode) {
      clearScheduleSelection();
    }
  });
  on(elements.scheduleSelectAllBtn, "click", selectAllVisibleScheduleItems);
  on(elements.scheduleExportSelectedBtn, "click", exportSelectedScheduleItems);
  on(elements.scheduleDaySelect, "change", async () => {
    const nextDate = elements.scheduleDaySelect.value;
    if (!nextDate) return;
    state.settings.app.scheduleView.activeDayKey = nextDate;
    state.settings.app.scheduleView.anchorDate = nextDate;
    await saveSettings();
    renderSchedule();
  });
  on(elements.scheduleDayPrev, "click", async () => {
    const current = elements.scheduleDaySelect?.value || todayKey();
    const nextDate = addDays(current, -1);
    state.settings.app.scheduleView.activeDayKey = nextDate;
    state.settings.app.scheduleView.anchorDate = nextDate;
    await saveSettings();
    renderSchedule();
  });
  on(elements.scheduleDayNext, "click", async () => {
    const current = elements.scheduleDaySelect?.value || todayKey();
    const nextDate = addDays(current, 1);
    state.settings.app.scheduleView.activeDayKey = nextDate;
    state.settings.app.scheduleView.anchorDate = nextDate;
    await saveSettings();
    renderSchedule();
  });
  on(elements.scheduleFiltersToggle, "click", (event) => {
    event.stopPropagation();
    const wrapper = elements.scheduleFiltersToggle.closest(".schedule-filters-wrapper");
    if (!wrapper) return;
    wrapper.classList.toggle("open");
    elements.scheduleFiltersToggle.setAttribute(
      "aria-expanded",
      String(wrapper.classList.contains("open"))
    );
  });
  document.addEventListener("click", (event) => {
    const wrapper = elements.scheduleFiltersToggle?.closest(".schedule-filters-wrapper");
    if (!wrapper || !wrapper.classList.contains("open")) return;
    if (wrapper.contains(event.target)) return;
    wrapper.classList.remove("open");
    elements.scheduleFiltersToggle?.setAttribute("aria-expanded", "false");
  });
  const updateToggleSettings = async () => {
    state.settings.app.scheduleView.toggles = {
      expectedOrders: elements.todayToggleExpected.checked,
      packs: elements.todayTogglePacks.checked,
      deliveries: elements.todayToggleDeliveries.checked,
    };
    await saveSettings();
    syncScheduleViewControls();
    renderAll();
  };
  on(elements.todayToggleExpected, "change", updateToggleSettings);
  on(elements.todayTogglePacks, "change", updateToggleSettings);
  on(elements.todayToggleDeliveries, "change", updateToggleSettings);
  on(elements.scheduleToggleExpected, "change", async () => {
    state.settings.app.scheduleView.toggles = {
      expectedOrders: elements.scheduleToggleExpected.checked,
      packs: elements.scheduleTogglePacks.checked,
      deliveries: elements.scheduleToggleDeliveries.checked,
    };
    await saveSettings();
    syncScheduleViewControls();
    renderAll();
  });
  on(elements.scheduleTogglePacks, "change", async () => {
    state.settings.app.scheduleView.toggles = {
      expectedOrders: elements.scheduleToggleExpected.checked,
      packs: elements.scheduleTogglePacks.checked,
      deliveries: elements.scheduleToggleDeliveries.checked,
    };
    await saveSettings();
    syncScheduleViewControls();
    renderAll();
  });
  on(elements.scheduleToggleDeliveries, "change", async () => {
    state.settings.app.scheduleView.toggles = {
      expectedOrders: elements.scheduleToggleExpected.checked,
      packs: elements.scheduleTogglePacks.checked,
      deliveries: elements.scheduleToggleDeliveries.checked,
    };
    await saveSettings();
    syncScheduleViewControls();
    renderAll();
  });

  const bindStatusChips = (container) => {
    if (!container) return;
    onDoc(`#${container.id} .filter-chip`, "click", async (_event, chip) => {
      state.settings.app.scheduleView.statusFilter = chip.dataset.status;
      await saveSettings();
      syncScheduleViewControls();
      renderAll();
    });
  };
  bindStatusChips(elements.scheduleStatusFilters);
  bindStatusChips(elements.todayStatusFilters);
  setupExportEvents();
}

async function init() {
  const appRoot = document.getElementById("appRoot");
  if (!appRoot) {
    console.warn("Missing #appRoot - UI not rendered yet");
  }
  setupGlobalErrorHandling();
  await initDB();
  await loadState();
  renderTabs();
  renderExport();
  setupEvents();
  setupAuthEvents();
  updateCloudStatusText();
  if (!supabaseAvailable) {
    console.error("Supabase client unavailable", supabaseInitError);
    showAppShell();
    setSyncError("Cloud sync unavailable. Running in local-only mode.");
    updateConnectionStatus({
      online: navigator.onLine,
      canWrite: true,
      status: "Offline/Local",
      email: "",
    });
    renderAll();
    return;
  }
  const { data } = await getSession();
  await handleSession(data?.session || null);
}

init().catch((error) => {
  showErrorScreen(error, "Initialization error");
});
