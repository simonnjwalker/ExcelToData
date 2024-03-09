## [1.0.5] - 2024-03-09
### Fixed
- No longer duplicates all rows were a DataSet has more than one table.
https://github.com/simonnjwalker/ExcelToData/issues/9


## [1.0.4] - 2024-01-15
### Added
- CSV reading/writing.
https://github.com/simonnjwalker/ExcelToData/issues/8


## [1.0.3] - 2023-05-01
### Fixed
- Now correctly skips blank cells in headers and in data.
https://github.com/simonnjwalker/ExcelToData/issues/6

### Added
- infers the worksheet name <T> when List<T> is the required output.
https://github.com/simonnjwalker/ExcelToData/issues/7


## [1.0.2] - 2022-12-01
### Changed
- added sheet indexes to ToDataTable().
https://github.com/simonnjwalker/ExcelToData/issues/5

### Fixed
- Fixed columns and allowed indexes for sheets
- fixed to check header row for merged, missing, and duplicated field names.
https://github.com/simonnjwalker/ExcelToData/issues/4


## [1.0.1] - 2022-09-01
### Fixed
- GetDefaults() now returns a clone not a reference to the internal settings object.
- Error where List<T> has non-CLR types.

### Added
- GetOptionsClone() which returns a clone not a reference.
- Conversion to-and-from Json as plain text in an XLSX column.

## [1.0.0] - 2022-08-01
### Added
- Initial release.
