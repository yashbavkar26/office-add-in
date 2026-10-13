const express = require("express");
const cors = require("cors");
const db = require("./db");

const app = express();
const PORT = process.env.PORT || 3001;

// Allow CORS from Office Add-in (localhost:3000) and other dev origins
app.use(
  cors({
    origin: "*",
    methods: ["GET", "POST", "OPTIONS"],
    allowedHeaders: ["Content-Type", "Authorization"],
  })
);

// Support JSON payloads and text/plain (used by navigator.sendBeacon)
app.use(express.json({ limit: "10mb" }));
app.use(express.text({ type: "text/plain", limit: "10mb" }));

// Middleware to parse beacon payloads if sent as string
app.use((req, res, next) => {
  if (typeof req.body === "string" && req.body.trim().startsWith("{")) {
    try {
      req.body = JSON.parse(req.body);
    } catch (e) {
      // Keep as string if parsing fails
    }
  }
  next();
});

// Health check endpoint
app.get("/api/status", (req, res) => {
  res.json({ status: "online", timestamp: new Date().toISOString() });
});

// Test connection & ensure database exists
app.post("/api/test-connection", async (req, res) => {
  const config = req.body;
  try {
    const result = await db.testAndEnsureDatabase(config);
    res.json({ success: true, message: result.message });
  } catch (error) {
    console.error("Connection error:", error);
    res.status(500).json({
      success: false,
      error: error.message || "Failed to connect to MySQL database.",
      details: error.code || "UNKNOWN_ERROR",
    });
  }
});

// Sync / verify table schema
app.post("/api/sync-schema", async (req, res) => {
  const { config, tableName, columns } = req.body;

  if (!columns || !Array.isArray(columns) || columns.length === 0) {
    return res.status(400).json({ success: false, error: "Columns array is required." });
  }

  try {
    const result = await db.ensureTableSchema(config, tableName, columns);
    res.json({ success: true, ...result });
  } catch (error) {
    console.error("Schema error:", error);
    res.status(500).json({
      success: false,
      error: error.message || "Failed to synchronize table schema.",
    });
  }
});

// Batch sync rows (upserts dirty rows, deletes removed/empty rows, prunes shifted rows)
app.post("/api/sync-rows", async (req, res) => {
  const { config, tableName, columns, rows, rowsToDelete, maxRowNum, clearAll } = req.body;

  try {
    let deletedCount = 0;

    // 1. If clearAll is true (all data rows deleted from sheet)
    if (clearAll) {
      const clearRes = await db.clearAllRows(config, tableName);
      return res.json({
        success: true,
        message: "Cleared all data records.",
        deletedCount: clearRes.clearedCount,
        upsertedCount: 0,
      });
    }

    // 2. Ensure table and columns exist before inserting
    if (columns && Array.isArray(columns) && columns.length > 0) {
      await db.ensureTableSchema(config, tableName, columns);
    }

    // 3. Delete explicitly removed or blanked-out rows
    if (rowsToDelete && Array.isArray(rowsToDelete) && rowsToDelete.length > 0) {
      const delRes = await db.deleteRows(config, tableName, rowsToDelete);
      deletedCount += delRes.deletedCount;
    }

    // 4. Prune rows beyond maxRowNum if table shrunk
    if (maxRowNum && typeof maxRowNum === "number" && maxRowNum >= 1) {
      const pruneRes = await db.pruneRowsBeyondMax(config, tableName, maxRowNum);
      deletedCount += pruneRes.prunedCount;
    }

    // 5. Upsert dirty rows with data
    let upsertRes = { affectedRows: 0, rowCount: 0 };
    if (rows && Array.isArray(rows) && rows.length > 0) {
      upsertRes = await db.upsertRows(config, tableName, columns, rows);
    }

    res.json({
      success: true,
      upsertedCount: upsertRes.rowCount || 0,
      deletedCount,
      affectedRows: (upsertRes.affectedRows || 0) + deletedCount,
    });
  } catch (error) {
    console.error("Sync rows error:", error);
    res.status(500).json({
      success: false,
      error: error.message || "Failed to synchronize rows into MySQL.",
    });
  }
});

// Beacon endpoint for window.beforeunload / workbook close
app.post("/api/sync-beacon", async (req, res) => {
  try {
    const payload = req.body;
    if (payload) {
      if (payload.clearAll) {
        await db.clearAllRows(payload.config, payload.tableName);
      } else {
        if (payload.columns && payload.columns.length > 0) {
          await db.ensureTableSchema(payload.config, payload.tableName, payload.columns);
        }
        if (payload.rowsToDelete && payload.rowsToDelete.length > 0) {
          await db.deleteRows(payload.config, payload.tableName, payload.rowsToDelete);
        }
        if (payload.maxRowNum) {
          await db.pruneRowsBeyondMax(payload.config, payload.tableName, payload.maxRowNum);
        }
        if (payload.rows && payload.rows.length > 0) {
          await db.upsertRows(payload.config, payload.tableName, payload.columns, payload.rows);
        }
      }
      console.log(`[Beacon] Flushed sync on close.`);
    }
    res.status(200).send("OK");
  } catch (error) {
    console.error("[Beacon Error]:", error);
    res.status(500).send(error.message);
  }
});

app.listen(PORT, () => {
  console.log(`====================================================`);
  console.log(` Excel to MySQL Sync Bridge Server running on port ${PORT}`);
  console.log(` API Endpoint: http://localhost:${PORT}/api/status`);
  console.log(`====================================================`);
});
