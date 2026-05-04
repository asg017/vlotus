//! Datetime extension wiring for vlotus.
//!
//! Thin shim around `lotus_datetime::register_on_sheet` plus a tag
//! predicate used by the auto-style branch in `ui::draw_grid`. The
//! heavy lifting (six handlers + ~40 functions) lives in the
//! `lotus-datetime` crate.
//!
//! Gated behind the `datetime` Cargo feature (default-on). Disable
//! with `cargo build -p vlotus --no-default-features` for a smaller
//! native binary.

use lotus_core::{RegistryError, Sheet};
use lotus_datetime::tags;

/// True if `tag` is one of the six datetime type tags shipped by
/// `lotus-datetime`. Used by `App::datetime_tag_for_cell` to decide
/// whether to apply the peach auto-style.
pub fn is_datetime_tag(tag: &str) -> bool {
    matches!(
        tag,
        tags::DATE | tags::TIME | tags::DATETIME | tags::ZONED | tags::TIMEZONE | tags::SPAN
    )
}

/// Register every datetime type and function on `sheet`. Idempotent
/// guard: a second call errors with `RegistryError::DuplicateTypeTag`.
/// Call once per `Sheet::new` site (currently `Store::recalculate`).
pub fn register(sheet: &mut Sheet) -> Result<(), RegistryError> {
    lotus_datetime::register_on_sheet(sheet)
}

#[cfg(test)]
mod tests {
    use super::*;
    use lotus_core::CellValue;

    #[test]
    fn is_datetime_tag_recognises_all_six() {
        assert!(is_datetime_tag("jdate"));
        assert!(is_datetime_tag("jtime"));
        assert!(is_datetime_tag("jdatetime"));
        assert!(is_datetime_tag("jzoned"));
        assert!(is_datetime_tag("jtimezone"));
        assert!(is_datetime_tag("jspan"));
    }

    #[test]
    fn is_datetime_tag_rejects_others() {
        assert!(!is_datetime_tag(""));
        assert!(!is_datetime_tag("hyperlink"));
        assert!(!is_datetime_tag("date")); // close but missing the j-prefix
        assert!(!is_datetime_tag("JDATE"));
    }

    #[test]
    fn register_makes_jdate_callable_in_a_formula() {
        let mut sheet = Sheet::new();
        register(&mut sheet).unwrap();
        sheet
            .set_cells(&[("A1".into(), "=DATE(2025, 4, 27)".into())])
            .unwrap();
        match sheet.get("A1") {
            CellValue::Custom(cv) => assert_eq!(cv.type_tag, tags::DATE),
            other => panic!("expected jdate Custom, got {other:?}"),
        }
    }

    #[test]
    fn parse_literal_claims_iso_date_in_full_pipeline() {
        let mut sheet = Sheet::new();
        register(&mut sheet).unwrap();
        sheet
            .set_cells(&[("A1".into(), "2025-04-27".into())])
            .unwrap();
        assert_eq!(sheet.type_tag("A1"), Some(tags::DATE));
    }

    #[test]
    fn double_register_errors() {
        let mut sheet = Sheet::new();
        register(&mut sheet).unwrap();
        let err = register(&mut sheet).unwrap_err();
        assert!(matches!(err, RegistryError::DuplicateTypeTag(_)));
    }
}
