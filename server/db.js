const mysql = require("mysql2/promise");

/**
 * Creates or gets a connection pool for the given configuration.
 */
function createPool(config) {
  return mysql.createPool({
    host: config.host || "localhost",
    port: parseInt(config.port, 10) || 3306,
    user: config.user || "root",
    password: config.password || "",
    database: config.database || "excel_sync_db",
    waitForConnections: true,
    connectionLimit: 10,
    queueLimit: 0,
    decimalNumbers: true,
  });
}

/**
 * Tests connection to MySQL and creates the target database if it doesn't already exist.
 */
async function testAndEnsureDatabase(config) {
  // First connect to MySQL without specifying a database to test server reachability
  const adminConn = await mysql.createConnection({
    host: config.host || "localhost",
    port: parseInt(config.port, 10) || 3306,
    user: config.user || "root",
    password: config.password || "",
  });

  try {
    const dbName = config.database || "excel_sync_db";
    // Sanitize database name (alphanumeric and underscores only)
    const sanitizedDbName = dbName.replace(/[^a-zA-Z0-9_]/g, "");
    if (!sanitizedDbName) {
      throw new Error("Invalid database name provided.");
    }

    await adminConn.query(`CREATE DATABASE IF NOT EXISTS \`${sanitizedDbName}\` CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci`);
    return { success: true, message: `Connected to MySQL. Database '${sanitizedDbName}' is ready.` };
  } finally {
    await adminConn.end();
  }
}

/**
 * Maps JS/Excel inferred data types to standard MySQL column data types.
 */
function mapToSqlType(type) {
  switch (type) {
    case "integer":
      return "INT";
    case "number":
    case "decimal":
      return "DOUBLE";
    case "boolean":
      return "TINYINT(1)";
    case "date":
    case "datetime":
      return "DATETIME";
    case "text":
      return "TEXT";
    case "string":
    default:
      return "VARCHAR(255)";
  }
}

/**
 * Sanitizes column or table names to prevent SQL injection and comply with MySQL identifiers.
 */
function sanitizeIdentifier(name) {
  return name.trim().replace(/[^a-zA-Z0-9_]/g, "_").toLowerCase();
}

/**
 * Ensures that the target table exists with the required columns.
 * Dynamically alters table if new columns are detected in Excel.
 */
async function ensureTableSchema(config, rawTableName, columns) {
  const pool = createPool(config);
  const tableName = sanitizeIdentifier(rawTableName || "orders");

  try {
    // Check if table exists
    const [existingTables] = await pool.query("SHOW TABLES LIKE ?", [tableName]);
    const tableExists = existingTables.length > 0;

    if (!tableExists) {
      // Build initial CREATE TABLE statement
      const colDefs = [
        "`_excel_row_num` INT NOT NULL UNIQUE",
      ];

      for (const col of columns) {
        const colName = sanitizeIdentifier(col.name);
        if (colName === "_excel_row_num") continue;

        const sqlType = mapToSqlType(col.type);
        if (col.isPrimary) {
          colDefs.push(`\`${colName}\` ${sqlType} NOT NULL PRIMARY KEY`);
        } else {
          colDefs.push(`\`${colName}\` ${sqlType} NULL`);
        }
      }

      colDefs.push("`_created_at` TIMESTAMP DEFAULT CURRENT_TIMESTAMP");
      colDefs.push("`_updated_at` TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP");

      const createSql = `CREATE TABLE \`${tableName}\` (${colDefs.join(", ")}) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;`;
      await pool.query(createSql);
      return { created: true, altered: false, tableName };
    } else {
      // Table exists: verify whether all columns exist, add missing ones
      const [existingCols] = await pool.query(`DESCRIBE \`${tableName}\``);
      const existingColNames = new Set(existingCols.map((c) => c.Field.toLowerCase()));

      let altered = false;
      for (const col of columns) {
        const colName = sanitizeIdentifier(col.name);
        if (colName === "_excel_row_num") continue;

        if (!existingColNames.has(colName)) {
          const sqlType = mapToSqlType(col.type);
          await pool.query(`ALTER TABLE \`${tableName}\` ADD COLUMN \`${colName}\` ${sqlType} NULL;`);
          altered = true;
        }
      }

      return { created: false, altered, tableName };
    }
  } finally {
    await pool.end();
  }
}

/**
 * Upserts a batch of rows into the MySQL database.
 * If the row already exists (identified by primary key or _excel_row_num),
 * updates the changed fields in-place so no duplicate row is created.
 */
async function upsertRows(config, rawTableName, columns, rows) {
  if (!rows || rows.length === 0) {
    return { affectedRows: 0, message: "No rows provided to sync." };
  }

  const pool = createPool(config);
  const tableName = sanitizeIdentifier(rawTableName || "orders");

  try {
    // Standardize column names
    const sanitizedCols = columns.map((c) => ({
      raw: c.name,
      clean: sanitizeIdentifier(c.name),
      isPrimary: !!c.isPrimary,
    }));

    // Ensure _excel_row_num is in column list
    const hasRowNumCol = sanitizedCols.some((c) => c.clean === "_excel_row_num");
    if (!hasRowNumCol) {
      sanitizedCols.unshift({ raw: "_excel_row_num", clean: "_excel_row_num", isPrimary: false });
    }

    const colNames = sanitizedCols.map((c) => `\`${c.clean}\``);

    // Build ON DUPLICATE KEY UPDATE clause
    // Update all data columns + _updated_at (clearing a cell sets that column to NULL in MySQL)
    const updateClauses = sanitizedCols
      .filter((c) => c.clean !== "_excel_row_num")
      .map((c) => `\`${c.clean}\` = VALUES(\`${c.clean}\`)`);

    updateClauses.push("`_updated_at` = CURRENT_TIMESTAMP");

    // Format row values array for mysql2
    const valuesMatrix = rows.map((row) => {
      return sanitizedCols.map((c) => {
        const val = row[c.clean] !== undefined ? row[c.clean] : row[c.raw];
        if (val === undefined || val === null || val === "") {
          return null;
        }
        return val;
      });
    });

    const placeholders = valuesMatrix
      .map(() => `(${new Array(sanitizedCols.length).fill("?").join(", ")})`)
      .join(", ");

    const flattenedValues = valuesMatrix.flat();

    const sql = `
      INSERT INTO \`${tableName}\` (${colNames.join(", ")})
      VALUES ${placeholders}
      ON DUPLICATE KEY UPDATE ${updateClauses.join(", ")};
    `;

    const [result] = await pool.query(sql, flattenedValues);
    return {
      success: true,
      affectedRows: result.affectedRows,
      tableName,
      rowCount: rows.length,
    };
  } finally {
    await pool.end();
  }
}

/**
 * Deletes specific rows by their Excel row number.
 */
async function deleteRows(config, rawTableName, rowNums) {
  if (!rowNums || !Array.isArray(rowNums) || rowNums.length === 0) {
    return { deletedCount: 0 };
  }

  const validRowNums = rowNums
    .map((n) => parseInt(n, 10))
    .filter((n) => !isNaN(n) && n > 0);

  if (validRowNums.length === 0) {
    return { deletedCount: 0 };
  }

  const pool = createPool(config);
  const tableName = sanitizeIdentifier(rawTableName || "orders");

  try {
    const placeholders = validRowNums.map(() => "?").join(", ");
    const sql = `DELETE FROM \`${tableName}\` WHERE \`_excel_row_num\` IN (${placeholders});`;
    const [result] = await pool.query(sql, validRowNums);
    return { deletedCount: result.affectedRows };
  } finally {
    await pool.end();
  }
}

/**
 * Prunes rows whose Excel row number is greater than maxRowNum.
 * When rows are deleted from the sheet, subsequent rows shift and the total count shrinks.
 */
async function pruneRowsBeyondMax(config, rawTableName, maxRowNum) {
  const max = parseInt(maxRowNum, 10);
  if (isNaN(max) || max < 1) {
    return { prunedCount: 0 };
  }

  const pool = createPool(config);
  const tableName = sanitizeIdentifier(rawTableName || "orders");

  try {
    const sql = `DELETE FROM \`${tableName}\` WHERE \`_excel_row_num\` > ?;`;
    const [result] = await pool.query(sql, [max]);
    return { prunedCount: result.affectedRows };
  } finally {
    await pool.end();
  }
}

/**
 * Clears all data rows from the table (rows >= 2).
 */
async function clearAllRows(config, rawTableName) {
  const pool = createPool(config);
  const tableName = sanitizeIdentifier(rawTableName || "orders");

  try {
    const sql = `DELETE FROM \`${tableName}\` WHERE \`_excel_row_num\` >= 2;`;
    const [result] = await pool.query(sql);
    return { clearedCount: result.affectedRows };
  } finally {
    await pool.end();
  }
}

module.exports = {
  testAndEnsureDatabase,
  ensureTableSchema,
  upsertRows,
  deleteRows,
  pruneRowsBeyondMax,
  clearAllRows,
  sanitizeIdentifier,
};
