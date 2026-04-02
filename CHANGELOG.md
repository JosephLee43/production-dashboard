# Changelog

All notable changes to the Production Dashboard project are documented in this file.

This project follows a Keep a Changelog style and uses dates for release tracking.

## [2026-04-02] - Dashboard reliability + UX updates

### Added
- Simulation mode toggle to switch between live and synthetic full-page data.
- Overdue visibility toggle to include/exclude past units from Work in Process view.
- Header filter-state badges (`Today+` / `Incl. Overdue`) for quick mode clarity.
- WIP visual indicators:
  - static green dot in WIP badge
  - smooth pulsing green header border for active WIP cards

### Changed
- Workstation allowlist fixed to:
  - `IXCWS1`
  - `IXCWS2`
  - `ZTSWS1`
  - `ZTSWS2`
  - `ZTSWS3`
- Work in Process and Completed rendering refined by status model:
  - WIP panel contains non-completed active items
  - Completed panel contains `COMPLETED` and `QC WAIT`
- Not Started behavior merged into Work in Process display logic.
- Date filtering and workstation normalization improved for real-world Excel values.
- WIP highlighting tuned for readability without causing scrollbar flicker.

### Fixed
- Excel parsing robustness for mixed cell value types:
  - date objects
  - formula result objects
  - rich text objects
- Station alias mapping compatibility (`ZTHWS1` -> `ZTSWS1`) to prevent hidden units.
- Auto-refresh reliability on workbook updates by hardening watcher logic:
  - queued refresh processing
  - `awaitWriteFinish` for OneDrive-style save patterns
  - multi-event watch coverage (`change`, `add`)
- Reduced missed client updates during rapid consecutive file save events.

### Notes
- Current runtime entrypoint is `server.js` on port `3000`.
- Dashboard serves from `public/` and updates clients via Socket.IO.
