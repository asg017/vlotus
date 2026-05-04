use rusqlite::Connection;

/// Magic identifier written to `PRAGMA application_id`. Spells "VLOT" in
/// ASCII; lets `file(1)` and other tools recognise a vlotus DB.
pub const APPLICATION_ID: i32 = 0x564C_4F54;

/// Schema version stored in `PRAGMA user_version`. Bump when the on-disk
/// shape of any user-data table changes in a way `ensure_schema` can't
/// idempotently bring forward.
pub const SCHEMA_VERSION: i32 = 1;

/// Composite-PK schema. Natural keys throughout so SQLite session-extension
/// changesets are content-addressable across DBs (see ticket coisl46s).
const DDL: &str = r#"
CREATE TABLE IF NOT EXISTS meta (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL
) WITHOUT ROWID;

CREATE TABLE IF NOT EXISTS sheet (
    name TEXT PRIMARY KEY,
    sort_order INTEGER NOT NULL DEFAULT 0,
    color TEXT,
    created_at TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ','now')),
    updated_at TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ','now'))
) WITHOUT ROWID;

CREATE TABLE IF NOT EXISTS column_meta (
    sheet_name TEXT NOT NULL
        REFERENCES sheet(name) ON UPDATE CASCADE ON DELETE CASCADE,
    col INTEGER NOT NULL,
    width INTEGER,
    PRIMARY KEY (sheet_name, col)
) WITHOUT ROWID;

CREATE TABLE IF NOT EXISTS cell (
    sheet_name TEXT NOT NULL
        REFERENCES sheet(name) ON UPDATE CASCADE ON DELETE CASCADE,
    row INTEGER NOT NULL,
    col INTEGER NOT NULL,
    raw TEXT NOT NULL DEFAULT '',
    computed TEXT,
    owner_row INTEGER,
    owner_col INTEGER,
    PRIMARY KEY (sheet_name, row, col)
) WITHOUT ROWID;

CREATE INDEX IF NOT EXISTS idx_cell_owner
    ON cell(sheet_name, owner_row, owner_col)
    WHERE owner_row IS NOT NULL;

CREATE TABLE IF NOT EXISTS cell_format (
    sheet_name TEXT NOT NULL,
    row INTEGER NOT NULL,
    col INTEGER NOT NULL,
    -- Opaque JSON blob today; the patch epic (att coisl46s) will
    -- consider re-typifying this so changesets carry per-axis diffs.
    format_json TEXT NOT NULL,
    PRIMARY KEY (sheet_name, row, col),
    FOREIGN KEY (sheet_name, row, col)
        REFERENCES cell(sheet_name, row, col)
        ON UPDATE CASCADE ON DELETE CASCADE
) WITHOUT ROWID;

CREATE TABLE IF NOT EXISTS undo_entry (
    id INTEGER PRIMARY KEY,
    group_id INTEGER NOT NULL,
    kind TEXT NOT NULL,
    payload TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_undo_group
    ON undo_entry(group_id);
"#;

pub fn ensure_schema(conn: &Connection) -> rusqlite::Result<()> {
    conn.execute_batch(DDL)
}

pub fn read_application_id(conn: &Connection) -> rusqlite::Result<i32> {
    conn.query_row("PRAGMA application_id", [], |r| r.get(0))
}

pub fn read_user_version(conn: &Connection) -> rusqlite::Result<i32> {
    conn.query_row("PRAGMA user_version", [], |r| r.get(0))
}

pub fn write_application_id(conn: &Connection, id: i32) -> rusqlite::Result<()> {
    conn.execute_batch(&format!("PRAGMA application_id = {id};"))
}

pub fn write_user_version(conn: &Connection, version: i32) -> rusqlite::Result<()> {
    conn.execute_batch(&format!("PRAGMA user_version = {version};"))
}

/// Returns true if the database has any pre-existing user table — used to
/// distinguish a genuinely empty fresh DB from one that just happens to
/// have its `application_id` and `user_version` defaulted to zero.
pub fn has_any_user_table(conn: &Connection) -> rusqlite::Result<bool> {
    conn.query_row(
        "SELECT EXISTS(\
            SELECT 1 FROM sqlite_master \
            WHERE type='table' AND name NOT LIKE 'sqlite_%')",
        [],
        |r| r.get::<_, bool>(0),
    )
}

/// Returns true if the database has a legacy `datasette_sheets_*` table.
/// A non-VLOT `application_id` plus this returning true is the signal to
/// surface `StoreError::LegacyDatabase` (and direct the user at the
/// future `vlotus migrate` subcommand, ticket cgmigrdy).
pub fn has_legacy_datasette_table(conn: &Connection) -> rusqlite::Result<bool> {
    conn.query_row(
        "SELECT EXISTS(\
            SELECT 1 FROM sqlite_master \
            WHERE type='table' AND name LIKE 'datasette_sheets_%')",
        [],
        |r| r.get::<_, bool>(0),
    )
}
