use std::path::PathBuf;

use thiserror::Error;

#[derive(Debug, Error)]
pub enum StoreError {
    #[error("sqlite error: {0}")]
    Sqlite(#[from] rusqlite::Error),

    #[error(
        "{path:?} is a legacy datasette_sheets database; \
         migrate it with `vlotus migrate <old> <new>` first"
    )]
    LegacyDatabase { path: PathBuf },

    #[error(
        "{path:?} is not a recognised vlotus database \
         (application_id={app_id:#010x}, user_version={user_version})"
    )]
    UnknownDatabase {
        path: PathBuf,
        app_id: i32,
        user_version: i32,
    },

    #[error(
        "{path:?} is a vlotus database from a future schema version \
         ({user_version}); this build supports up to v{supported}"
    )]
    UnsupportedSchemaVersion {
        path: PathBuf,
        user_version: i32,
        supported: i32,
    },

    #[error("formula engine: {0}")]
    Engine(String),

    #[error("io: {0}")]
    Io(#[from] std::io::Error),

    #[error("a patch is already being recorded into {path:?}")]
    PatchAlreadyOpen { path: PathBuf },

    #[error("no active patch")]
    NoActivePatch,

    #[error("sheet '{0}' already exists")]
    DuplicateSheetName(String),
}
