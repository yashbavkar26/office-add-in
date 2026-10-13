/* global console, document, window, Excel, Office, fetch, navigator, localStorage, RequestInit, Response, AbortController, setTimeout, clearTimeout */

interface DbConfig {
  host: string;
  port: number;
  database: string;
  table: string;
  user: string;
  password: string;
}

interface ColumnSchema {
  name: string;
  type: string;
  isPrimary?: boolean;
}

interface LogItem {
  time: string;
  msg: string;
  type: "info" | "success" | "warning" | "error";
}

interface DOMCache {
  sideloadMsg: HTMLElement | null;
  appBody: HTMLElement | null;
  connPill: HTMLElement | null;
  connText: HTMLElement | null;
  headerLastSync: HTMLElement | null;
  dirtyCount: HTMLElement | null;
  activeRowIndicator: HTMLElement | null;
  countdownTimer: HTMLElement | null;
  countdownProgress: HTMLElement | null;
  countdownPercentage: HTMLElement | null;
  autosyncState: HTMLElement | null;
  bridgeCardStatus: HTMLElement | null;
  bridgeCardSub: HTMLElement | null;
  bridgeStatus: HTMLElement | null;
  lastSyncTime: HTMLElement | null;
  activityLog: HTMLElement | null;
  cfgHost: HTMLInputElement | null;
  cfgPort: HTMLInputElement | null;
  cfgDb: HTMLInputElement | null;
  cfgTable: HTMLInputElement | null;
  cfgUser: HTMLInputElement | null;
  cfgPassword: HTMLInputElement | null;
  testFeedback: HTMLElement | null;
  dbSettingsToggle: HTMLElement | null;
  dbSettingsContent: HTMLElement | null;
  dbSettingsChevron: HTMLElement | null;
  cardDbSettings: HTMLElement | null;
}

// Server endpoint
const BACKEND_URL = "http://localhost:3001";
const SYNC_INTERVAL_SECONDS = 60;
const SYNC_INTERVAL_MS = SYNC_INTERVAL_SECONDS * 1000;
const MAX_LOG_ENTRIES = 50;

// State trackers
let dom: DOMCache;
let isPaneVisible = true;
let currentActiveRow: number | null = null;
const pendingRowEdits = new Set<number>();
const dirtyRowQueue = new Set<number>();
const recentLogs: LogItem[] = [];

let isAutoSyncEnabled = true;
let isSyncing = false;
let countdownTimer: any = null;
let lastSyncedDate: Date | null = null;

/**
 * Resilient fetch with timeout to prevent hung requests in production
 */
async function fetchWithTimeout(url: string, options: RequestInit = {}, timeoutMs = 6000): Promise<Response> {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    return await fetch(url, { ...options, signal: controller.signal });
  } finally {
    clearTimeout(timer);
  }
}

/**
 * Gets or initializes the target timestamp for next auto-sync from localStorage
 */
function getNextSyncTimestamp(): number {
  const saved = localStorage.getItem("excel_mysql_next_sync_target");
  const now = Date.now();
  if (saved) {
    const ts = parseInt(saved, 10);
    if (!isNaN(ts) && ts > now && ts <= now + SYNC_INTERVAL_MS + 5000) {
      return ts;
    }
  }
  const newTarget = now + SYNC_INTERVAL_MS;
  localStorage.setItem("excel_mysql_next_sync_target", newTarget.toString());
  return newTarget;
}

function resetNextSyncTimestamp(): number {
  const target = Date.now() + SYNC_INTERVAL_MS;
  localStorage.setItem("excel_mysql_next_sync_target", target.toString());
  updateTimerUI();
  return target;
}

let nextSyncTimestamp = getNextSyncTimestamp();

/**
 * Initialize Add-in when Office runtime is ready
 */
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    dom = cacheDOMElements();
    if (dom.sideloadMsg) dom.sideloadMsg.style.display = "none";
    if (dom.appBody) dom.appBody.style.display = "flex";

    initApp();
  }
});

function cacheDOMElements(): DOMCache {
  return {
    sideloadMsg: document.getElementById("sideload-msg"),
    appBody: document.getElementById("app-body"),
    connPill: document.getElementById("connection-status-pill"),
    connText: document.getElementById("connection-status-text"),
    headerLastSync: document.getElementById("header-last-sync"),
    dirtyCount: document.getElementById("dirty-count"),
    activeRowIndicator: document.getElementById("active-row-indicator"),
    countdownTimer: document.getElementById("countdown-timer"),
    countdownProgress: document.getElementById("countdown-progress"),
    countdownPercentage: document.getElementById("countdown-percentage"),
    autosyncState: document.getElementById("autosync-state"),
    bridgeCardStatus: document.getElementById("bridge-card-status"),
    bridgeCardSub: document.getElementById("bridge-card-sub"),
    bridgeStatus: document.getElementById("bridge-status"),
    lastSyncTime: document.getElementById("last-sync-time"),
    activityLog: document.getElementById("activity-log"),
    cfgHost: document.getElementById("cfg-host") as HTMLInputElement,
    cfgPort: document.getElementById("cfg-port") as HTMLInputElement,
    cfgDb: document.getElementById("cfg-db") as HTMLInputElement,
    cfgTable: document.getElementById("cfg-table") as HTMLInputElement,
    cfgUser: document.getElementById("cfg-user") as HTMLInputElement,
    cfgPassword: document.getElementById("cfg-password") as HTMLInputElement,
    testFeedback: document.getElementById("test-feedback"),
    dbSettingsToggle: document.getElementById("db-settings-toggle"),
    dbSettingsContent: document.getElementById("db-settings-content"),
    dbSettingsChevron: document.getElementById("db-settings-chevron"),
    cardDbSettings: document.getElementById("card-db-settings"),
  };
}

function initApp() {
  loadSavedSettings();
  setupUIEventListeners();
  setupExcelEventListeners();
  setupSharedRuntimeVisibility();
  checkBackendHealth();
  startCountdownTimer();
  setupUnloadFlush();
}

/**
 * Loads saved MySQL credentials from localStorage
 */
function loadSavedSettings() {
  const saved = localStorage.getItem("excel_mysql_sync_config");
  if (saved) {
    try {
      const cfg: DbConfig = JSON.parse(saved);
      if (dom.cfgHost) dom.cfgHost.value = cfg.host || "localhost";
      if (dom.cfgPort) dom.cfgPort.value = (cfg.port || 3306).toString();
      if (dom.cfgDb) dom.cfgDb.value = cfg.database || "excel_sync_db";
      if (dom.cfgTable) dom.cfgTable.value = cfg.table || "orders";
      if (dom.cfgUser) dom.cfgUser.value = cfg.user || "root";
      if (dom.cfgPassword) dom.cfgPassword.value = cfg.password || "";
    } catch (e) {
      console.warn("Failed to parse saved config", e);
    }
  }
}

/**
 * Retrieves current MySQL configuration from UI
 */
function getDbConfig(): DbConfig {
  return {
    host: dom.cfgHost?.value.trim() || "localhost",
    port: parseInt(dom.cfgPort?.value.trim() || "3306", 10) || 3306,
    database: dom.cfgDb?.value.trim() || "excel_sync_db",
    table: dom.cfgTable?.value.trim() || "orders",
    user: dom.cfgUser?.value.trim() || "root",
    password: dom.cfgPassword?.value || "",
  };
}

function saveConfig(cfg: DbConfig) {
  try {
    localStorage.setItem("excel_mysql_sync_config", JSON.stringify(cfg));
  } catch (e) {
    console.warn("Failed to persist config to localStorage", e);
  }
}

/**
 * Checks if the local Sync Bridge Server is running
 */
async function checkBackendHealth() {
  try {
    const res = await fetchWithTimeout(`${BACKEND_URL}/api/status`, {}, 4000);
    if (res.ok) {
      if (dom.bridgeStatus) dom.bridgeStatus.innerText = "Online (Port 3001)";
      if (dom.bridgeCardStatus) dom.bridgeCardStatus.innerText = "Online";
      if (dom.bridgeCardSub) dom.bridgeCardSub.innerText = "Port 3001";
      addLog("Sync Bridge Server online.", "success");
      testDbConnection(false);
    } else {
      throw new Error(`HTTP ${res.status}`);
    }
  } catch (err: any) {
    if (dom.bridgeStatus) dom.bridgeStatus.innerText = "Offline";
    if (dom.bridgeCardStatus) {
      dom.bridgeCardStatus.innerText = "Offline";
      dom.bridgeCardStatus.className = "kpi-value text-secondary";
    }
    if (dom.bridgeCardSub) dom.bridgeCardSub.innerText = "Run npm run server";
    if (dom.connPill) dom.connPill.className = "status-badge status-error";
    if (dom.connText) dom.connText.innerText = "Server Offline";
    addLog("Sync Bridge server not reachable at localhost:3001. Run 'npm run server'.", "error");
  }
}

/**
 * Tests connection to MySQL database
 */
async function testDbConnection(showVerboseLog = true) {
  const cfg = getDbConfig();
  if (dom.testFeedback) {
    dom.testFeedback.className = "feedback-banner";
    dom.testFeedback.innerText = "Testing connection...";
    dom.testFeedback.style.display = "block";
  }

  try {
    const res = await fetchWithTimeout(`${BACKEND_URL}/api/test-connection`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(cfg),
    }, 6000);

    const data = await res.json();
    if (data.success) {
      saveConfig(cfg);
      if (dom.connPill) dom.connPill.className = "status-badge status-connected";
      if (dom.connText) dom.connText.innerText = "Connected";
      if (dom.testFeedback) {
        dom.testFeedback.className = "feedback-banner success";
        dom.testFeedback.innerText = "Connected to MySQL successfully!";
      }
      if (showVerboseLog) {
        addLog(`MySQL Connected: ${cfg.user}@${cfg.host}:${cfg.port}/${cfg.database}`, "success");
      }
    } else {
      throw new Error(data.error || "Connection failed");
    }
  } catch (err: any) {
    if (dom.connPill) dom.connPill.className = "status-badge status-error";
    if (dom.connText) dom.connText.innerText = "DB Error";
    if (dom.testFeedback) {
      dom.testFeedback.className = "feedback-banner error";
      dom.testFeedback.innerText = `Error: ${err.message}`;
    }
    if (showVerboseLog) {
      addLog(`MySQL Connection error: ${err.message}`, "error");
    }
  }
}

/**
 * Setup UI DOM event listeners
 */
function setupUIEventListeners() {
  document.getElementById("btn-test-db")?.addEventListener("click", () => {
    testDbConnection(true);
  });

  document.getElementById("btn-sample-data")?.addEventListener("click", () => {
    populateSampleOrdersTable();
  });

  document.getElementById("btn-sync-now")?.addEventListener("click", () => {
    addLog("Manual sync triggered.", "info");
    flushDirtyQueue(true);
  });

  const toggle = document.getElementById("toggle-autosync") as HTMLInputElement;
  toggle?.addEventListener("change", (e) => {
    isAutoSyncEnabled = (e.target as HTMLInputElement).checked;
    if (dom.autosyncState) {
      dom.autosyncState.innerText = isAutoSyncEnabled ? "Running (60s cycle)" : "Paused";
    }
    addLog(`Auto-sync ${isAutoSyncEnabled ? "resumed" : "paused"}.`, "info");
  });

  dom.dbSettingsToggle?.addEventListener("click", () => {
    if (dom.dbSettingsContent) {
      const isHidden = dom.dbSettingsContent.style.display === "none";
      dom.dbSettingsContent.style.display = isHidden ? "flex" : "none";
      if (dom.dbSettingsChevron) {
        dom.dbSettingsChevron.style.transform = isHidden ? "rotate(0deg)" : "rotate(-90deg)";
      }
    }
  });

  document.getElementById("btn-quick-settings")?.addEventListener("click", () => {
    if (dom.dbSettingsContent) {
      dom.dbSettingsContent.style.display = "flex";
      if (dom.dbSettingsChevron) dom.dbSettingsChevron.style.transform = "rotate(0deg)";
      dom.cardDbSettings?.scrollIntoView({ behavior: "smooth" });
    }
  });

  const togglePwBtn = document.getElementById("toggle-password-visibility");
  togglePwBtn?.addEventListener("click", () => {
    if (dom.cfgPassword) {
      const isPassword = dom.cfgPassword.type === "password";
      dom.cfgPassword.type = isPassword ? "text" : "password";
      togglePwBtn.style.color = isPassword ? "var(--primary)" : "var(--text-secondary)";
    }
  });

  document.getElementById("btn-clear-log")?.addEventListener("click", () => {
    recentLogs.length = 0;
    if (dom.activityLog) dom.activityLog.innerHTML = "";
  });
}

/**
 * Setup Excel Sheet event listeners
 */
async function setupExcelEventListeners() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.onChanged.add(handleWorksheetChanged);
      sheet.onSelectionChanged.add(handleSelectionChanged);
      await context.sync();
      addLog("Excel event listeners registered (change & selection).", "info");
    });
  } catch (err: any) {
    addLog(`Failed to attach Excel listeners: ${err.message}`, "error");
  }
}

/**
 * Handles cell value edits in the worksheet
 */
async function handleWorksheetChanged(event: Excel.WorkSheetChangedEventArgs) {
  try {
    if (event.changeType === "RowDeleted" || event.changeType === "CellDeleted") {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const usedRange = sheet.getUsedRangeOrNullObject();
        usedRange.load(["rowCount", "rowIndex", "isNullObject"]);
        await context.sync();

        if (usedRange.isNullObject || usedRange.rowCount <= 1) {
          dirtyRowQueue.add(2);
        } else {
          const startRow = usedRange.rowIndex + 1;
          const totalRows = usedRange.rowCount;
          for (let r = 2; r <= startRow + totalRows; r++) {
            dirtyRowQueue.add(r);
          }
        }
        updateUIState();
        addLog(`[Delete] Detected deletion in Excel. Queued table sync.`, "warning");
      });
      return;
    }

    await Excel.run(async (context) => {
      const range = context.workbook.worksheets.getActiveWorksheet().getRange(event.address);
      range.load(["rowIndex", "rowCount"]);
      await context.sync();

      const startRow = range.rowIndex + 1;
      const rowCount = range.rowCount;

      for (let i = 0; i < rowCount; i++) {
        const rowNum = startRow + i;
        if (rowNum === 1) continue; // Skip header

        if (currentActiveRow !== null && rowNum === currentActiveRow) {
          pendingRowEdits.add(rowNum);
          updateUIState();
          addLog(`Editing Row #${rowNum}... (held in active state)`, "warning");
        } else {
          dirtyRowQueue.add(rowNum);
          updateUIState();
          addLog(`Row #${rowNum} modified and added to Dirty Row Queue.`, "info");
        }
      }
    });
  } catch (e: any) {
    console.error("Error in handleWorksheetChanged:", e);
  }
}

/**
 * Handles cursor navigation / row exit
 */
async function handleSelectionChanged(event: Excel.WorkSheetSelectionChangedEventArgs) {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.worksheets.getActiveWorksheet().getRange(event.address);
      range.load(["rowIndex"]);
      await context.sync();

      const newRow = range.rowIndex + 1;

      if (currentActiveRow !== null && currentActiveRow !== newRow) {
        if (pendingRowEdits.has(currentActiveRow)) {
          pendingRowEdits.delete(currentActiveRow);
          dirtyRowQueue.add(currentActiveRow);
          addLog(`Finished editing Row #${currentActiveRow} → Queued (${dirtyRowQueue.size} in queue).`, "info");
        }
      }

      currentActiveRow = newRow;
      updateUIState();
    });
  } catch (e: any) {
    console.error("Error in handleSelectionChanged:", e);
  }
}

/**
 * Updates countdown UI elements based on exact wall-clock remaining seconds
 */
function updateTimerUI() {
  if (!isPaneVisible) return; // Skip DOM updates when taskpane is hidden

  const now = Date.now();
  const remainingMs = Math.max(0, nextSyncTimestamp - now);
  const remainingSeconds = Math.ceil(remainingMs / 1000);

  if (dom.countdownTimer) {
    dom.countdownTimer.innerText = `${remainingSeconds}s`;
  }
  if (dom.countdownProgress) {
    const percentage = Math.min(100, Math.max(0, (remainingSeconds / SYNC_INTERVAL_SECONDS) * 100));
    dom.countdownProgress.style.width = `${percentage}%`;
    if (dom.countdownPercentage) {
      dom.countdownPercentage.innerText = `${Math.round(percentage)}%`;
    }
  }
}

/**
 * Handles Shared Runtime visibility changes: saves CPU by pausing DOM updates
 */
function setupSharedRuntimeVisibility() {
  if (Office.addin && (Office.addin as any).onVisibilityModeChanged) {
    (Office.addin as any).onVisibilityModeChanged((args: any) => {
      isPaneVisible = (args.visibilityMode === (Office as any).VisibilityMode?.taskpane || args.visibilityMode === "Taskpane");
      if (isPaneVisible) {
        renderRecentLogs();
        updateUIState();
        updateTimerUI();
        updateLastSyncTime();
        addLog("Task pane reopened. Background sync active.", "info");
      } else {
        addLog("Task pane hidden. Auto-sync continuing in background.", "info");
      }
    });
  }
}

/**
 * Persistent countdown timer ticker based on wall-clock time
 */
function startCountdownTimer() {
  if (countdownTimer) clearInterval(countdownTimer);

  updateTimerUI();

  countdownTimer = setInterval(() => {
    if (!isAutoSyncEnabled) return;

    const now = Date.now();
    const remainingMs = nextSyncTimestamp - now;

    if (remainingMs <= 0) {
      nextSyncTimestamp = resetNextSyncTimestamp();
      updateTimerUI();
      flushDirtyQueue(false);
    } else {
      updateTimerUI();
    }
  }, 1000);
}

/**
 * Flushes dirty rows to MySQL
 */
async function flushDirtyQueue(isManual: boolean) {
  if (isSyncing) return;

  const rowsToSync = Array.from(dirtyRowQueue);

  if (isManual && currentActiveRow !== null && pendingRowEdits.has(currentActiveRow)) {
    pendingRowEdits.delete(currentActiveRow);
    dirtyRowQueue.add(currentActiveRow);
    rowsToSync.push(currentActiveRow);
  }

  if (rowsToSync.length === 0) {
    if (pendingRowEdits.size > 0) {
      addLog(`[Auto-Sync] Row #${Array.from(pendingRowEdits).join(", #")} is still being edited.`, "info");
    }
    return;
  }

  isSyncing = true;
  updateUIState();
  addLog(`[Sync] Checking worksheet changes for ${rowsToSync.length} row(s)...`, "info");

  try {
    const cfg = getDbConfig();

    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = sheet.getUsedRangeOrNullObject();
      usedRange.load(["values", "rowCount", "columnCount", "isNullObject"]);
      await context.sync();

      if (usedRange.isNullObject || !usedRange.values || usedRange.values.length <= 1) {
        addLog("Worksheet has no data rows. Pruning all records from MySQL...", "warning");
        const res = await fetchWithTimeout(`${BACKEND_URL}/api/sync-rows`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            config: cfg,
            tableName: cfg.table,
            clearAll: true,
          }),
        }, 8000);

        const result = await res.json();
        if (result.success) {
          dirtyRowQueue.clear();
          pendingRowEdits.clear();
          lastSyncedDate = new Date();
          updateLastSyncTime();
          nextSyncTimestamp = resetNextSyncTimestamp();
          addLog(`Cleared all records from MySQL table '${cfg.table}'.`, "success");
        }
        return;
      }

      const values = usedRange.values;
      const rawHeaders = values[0];
      const columns: ColumnSchema[] = [];

      for (let c = 0; c < rawHeaders.length; c++) {
        const headerName = String(rawHeaders[c] || `Col_${c + 1}`).trim();
        let inferredType = "string";
        for (let r = 1; r < Math.min(values.length, 10); r++) {
          const cellVal = values[r][c];
          if (cellVal !== null && cellVal !== undefined && cellVal !== "") {
            if (typeof cellVal === "number") {
              inferredType = Number.isInteger(cellVal) ? "integer" : "decimal";
            } else if (typeof cellVal === "boolean") {
              inferredType = "boolean";
            } else if (!isNaN(Date.parse(cellVal)) && isNaN(Number(cellVal))) {
              inferredType = "datetime";
            }
            break;
          }
        }
        columns.push({ name: headerName, type: inferredType, isPrimary: false });
      }

      const payloadRows: any[] = [];
      const rowsToDelete: number[] = [];
      const successfullyPreparedRows: number[] = [];

      for (const rowNum of rowsToSync) {
        const rowIndex = rowNum - 1;
        if (rowIndex < values.length) {
          const rowValues = values[rowIndex];
          const rowObj: any = { _excel_row_num: rowNum };

          let hasAnyData = false;
          for (let c = 0; c < columns.length; c++) {
            const val = rowValues[c] !== undefined ? rowValues[c] : null;
            rowObj[columns[c].name] = val;
            if (val !== null && val !== undefined && String(val).trim() !== "") {
              hasAnyData = true;
            }
          }

          if (hasAnyData) {
            payloadRows.push(rowObj);
          } else {
            rowsToDelete.push(rowNum);
          }
          successfullyPreparedRows.push(rowNum);
        } else {
          rowsToDelete.push(rowNum);
          successfullyPreparedRows.push(rowNum);
        }
      }

      const maxRowNum = values.length;

      const res = await fetchWithTimeout(`${BACKEND_URL}/api/sync-rows`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          config: cfg,
          tableName: cfg.table,
          columns,
          rows: payloadRows,
          rowsToDelete,
          maxRowNum,
        }),
      }, 10000);

      const result = await res.json();
      if (result.success) {
        for (const syncedRow of successfullyPreparedRows) {
          dirtyRowQueue.delete(syncedRow);
        }

        lastSyncedDate = new Date();
        updateLastSyncTime();
        nextSyncTimestamp = resetNextSyncTimestamp();

        const summaryParts: string[] = [];
        if (payloadRows.length > 0) {
          summaryParts.push(`Synced ${payloadRows.length} active row(s)`);
        }
        if (result.deletedCount && result.deletedCount > 0) {
          summaryParts.push(`Deleted ${result.deletedCount} removed record(s)`);
        }
        if (summaryParts.length === 0) {
          summaryParts.push("Database is up to date");
        }

        addLog(`${summaryParts.join(" and ")} in MySQL table '${cfg.table}'.`, "success");
      } else {
        throw new Error(result.error || "Failed to sync rows.");
      }
    });
  } catch (err: any) {
    addLog(`Sync error: ${err.message}`, "error");
  } finally {
    isSyncing = false;
    updateUIState();
  }
}

/**
 * Flush on workbook/window close using Beacon API
 */
function setupUnloadFlush() {
  window.addEventListener("beforeunload", () => {
    const rowsToFlush = Array.from(new Set([...dirtyRowQueue, ...pendingRowEdits]));
    if (rowsToFlush.length === 0) return;

    const cfg = getDbConfig();
    try {
      Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const usedRange = sheet.getUsedRangeOrNullObject();
        usedRange.load(["values", "isNullObject"]);
        await context.sync();

        if (usedRange.isNullObject || !usedRange.values || usedRange.values.length <= 1) {
          navigator.sendBeacon(`${BACKEND_URL}/api/sync-beacon`, JSON.stringify({ config: cfg, tableName: cfg.table, clearAll: true }));
          return;
        }

        const values = usedRange.values;
        const rawHeaders = values[0];
        const columns = rawHeaders.map((h: any) => ({ name: String(h || "col"), type: "string", isPrimary: false }));
        const rows: any[] = [];
        const rowsToDelete: number[] = [];

        for (const rowNum of rowsToFlush) {
          const idx = rowNum - 1;
          if (idx < values.length) {
            const rowObj: any = { _excel_row_num: rowNum };
            let hasAny = false;
            for (let c = 0; c < columns.length; c++) {
              const val = values[idx][c];
              rowObj[columns[c].name] = val;
              if (val !== null && val !== undefined && String(val).trim() !== "") {
                hasAny = true;
              }
            }
            if (hasAny) {
              rows.push(rowObj);
            } else {
              rowsToDelete.push(rowNum);
            }
          } else {
            rowsToDelete.push(rowNum);
          }
        }

        const payload = JSON.stringify({
          config: cfg,
          tableName: cfg.table,
          columns,
          rows,
          rowsToDelete,
          maxRowNum: values.length,
        });
        navigator.sendBeacon(`${BACKEND_URL}/api/sync-beacon`, payload);
      });
    } catch (e) {
      console.error("Failed beforeunload beacon flush", e);
    }
  });
}

/**
 * Updates UI counter badges and indicators
 */
function updateUIState() {
  if (!isPaneVisible) return;

  if (dom.dirtyCount) {
    dom.dirtyCount.innerText = dirtyRowQueue.size.toString();
  }

  if (dom.activeRowIndicator) {
    if (currentActiveRow !== null && pendingRowEdits.has(currentActiveRow)) {
      dom.activeRowIndicator.innerText = `Editing Row #${currentActiveRow}`;
      dom.activeRowIndicator.style.color = "var(--warning)";
    } else if (currentActiveRow !== null) {
      dom.activeRowIndicator.innerText = `Active on Row #${currentActiveRow}`;
      dom.activeRowIndicator.style.color = "var(--text-secondary)";
    } else {
      dom.activeRowIndicator.innerText = "No row active";
      dom.activeRowIndicator.style.color = "var(--text-secondary)";
    }
  }
}

/**
 * Populates sample dummy Orders table in Excel
 */
async function populateSampleOrdersTable() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.getUsedRange().clear(Excel.ClearApplyTo.all);

      const data = [
        ["order_id", "customer_name", "product", "category", "quantity", "unit_price", "total_amount", "order_status", "order_date"],
        ["ORD-1001", "Alice Johnson", "Wireless Mechanical Keyboard", "Electronics", 2, 89.99, 179.98, "Delivered", "2026-09-15"],
        ["ORD-1002", "Brian Smith", "Ergonomic Office Chair", "Furniture", 1, 249.50, 249.50, "Shipped", "2026-09-18"],
        ["ORD-1003", "Carla Gomez", "USB-C Multi-Port Hub", "Accessories", 3, 34.00, 102.00, "Processing", "2026-09-22"],
        ["ORD-1004", "David Lee", "Ultra-Wide Monitor 34-inch", "Electronics", 1, 499.99, 499.99, "Delivered", "2026-09-25"],
        ["ORD-1005", "Emma Watson", "Noise Cancelling Headphones", "Audio", 2, 129.00, 258.00, "Pending", "2026-09-28"],
      ];

      const range = sheet.getRange("A1:I6");
      range.values = data;

      const headerRange = sheet.getRange("A1:I1");
      headerRange.format.fill.color = "#0F6CBD";
      headerRange.format.font.bold = true;
      headerRange.format.font.color = "#FFFFFF";

      range.format.autofitColumns();
      await context.sync();

      addLog("Sample Orders table inserted into Excel.", "success");
    });
  } catch (err: any) {
    addLog(`Error inserting sample data: ${err.message}`, "error");
  }
}

/**
 * Updates last sync timestamp in the UI
 */
function updateLastSyncTime() {
  if (lastSyncedDate) {
    const formatted = lastSyncedDate.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit", second: "2-digit" });
    if (dom.lastSyncTime) dom.lastSyncTime.innerText = `Last Sync: ${formatted}`;
    if (dom.headerLastSync) dom.headerLastSync.innerText = `Last sync: ${formatted}`;
  }
}

/**
 * Creates structured DOM element for audit log entry
 */
function createLogElement(item: LogItem): HTMLElement {
  const row = document.createElement("div");
  row.className = `log-row ${item.type}`;

  const ts = document.createElement("span");
  ts.className = "log-time";
  ts.innerText = item.time;

  const badge = document.createElement("span");
  badge.className = `badge badge-${item.type}`;
  badge.innerText = item.type === "success" ? "OK" : item.type === "warning" ? "WARN" : item.type === "error" ? "ERR" : "INFO";

  const text = document.createElement("span");
  text.className = "log-message";
  text.innerText = item.msg;

  row.appendChild(ts);
  row.appendChild(badge);
  row.appendChild(text);
  return row;
}

/**
 * Replays recent buffered logs into the activity-log element
 */
function renderRecentLogs() {
  if (!dom.activityLog) return;
  dom.activityLog.innerHTML = "";

  const fragment = document.createDocumentFragment();
  for (const item of recentLogs) {
    fragment.appendChild(createLogElement(item));
  }
  dom.activityLog.appendChild(fragment);
  dom.activityLog.scrollTop = dom.activityLog.scrollHeight;
}

/**
 * Appends message to the live activity console log and memory buffer
 */
function addLog(msg: string, type: "info" | "success" | "warning" | "error" = "info") {
  const time = new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit", second: "2-digit" });
  const logItem: LogItem = { time, msg, type };
  recentLogs.push(logItem);
  if (recentLogs.length > MAX_LOG_ENTRIES) recentLogs.shift();

  if (!isPaneVisible || !dom.activityLog) return;

  dom.activityLog.appendChild(createLogElement(logItem));
  if (dom.activityLog.childElementCount > MAX_LOG_ENTRIES) {
    dom.activityLog.removeChild(dom.activityLog.firstElementChild!);
  }
  dom.activityLog.scrollTop = dom.activityLog.scrollHeight;
}
