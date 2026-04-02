# Changelog

All notable changes to the Production Dashboard project are documented in this file.

This project follows a Keep a Changelog style and uses dates for release tracking.

---

## [2026-04-02] — Dashboard docs + screenshot

### Added
- Dashboard screenshot added to `docs/images/dashboard-example.jpg`.
- `README.md` updated with screenshot section and project setup instructions.
- `docs/images/` directory created for project documentation assets.

---

## [2026-04-02] — Dashboard reliability + UX updates

### Added
- Simulation mode toggle to switch between live and synthetic full-page data.
- Overdue visibility toggle to include/exclude past units from Work in Process view.
- Header filter-state badges (`Today+` / `Incl. Overdue`) for quick mode clarity.
- WIP visual indicators:
  - Static green dot in WIP status badge.
  - Smooth pulsing green inset header border for active (in-progress) WIP cards (2.4s cubic-bezier).
  - Shared `--wip-accent-rgb` CSS variable drives both dot and border colour.

### Changed
- Workstation allowlist fixed to: `IXCWS1`, `IXCWS2`, `ZTSWS1`, `ZTSWS2`, `ZTSWS3`.
- Work in Process and Completed rendering refined by status:
  - WIP panel: all non-completed active items (`NOT STARTED`, `WIP`, `BUILD COMPLETE`, etc.).
  - Completed panel: `COMPLETED` and `QC WAIT` only.
- Not Started cards merged into WIP panel with distinct visual treatment (no animated border).
- Date filtering and workstation normalisation improved for real-world Excel data.
- WIP animation switched from `transform: scale()` to `inset box-shadow` to prevent scrollbar flicker.

### Fixed
- Excel parsing robustness for mixed cell value types (date objects, formula results, rich text objects).
- Station alias mapping: `ZTHWS1` -> `ZTSWS1` so units are not silently dropped.
- Auto-refresh reliability when Excel file is saved via OneDrive or in rapid succession:
  - Queued refresh worker prevents duplicate parallel reads.
  - `awaitWriteFinish: { stabilityThreshold: 1800, pollInterval: 150 }` added to chokidar.
  - Multi-event watch coverage (`change` and `add` both trigger refresh).

### Notes
- Runtime entrypoint: `server.js` on port `3000`.
- Frontend served from `public/`, real-time updates via Socket.IO.

---

## [2026-04-02] — Initial project creation

### Added
- `server.js`: Express + Socket.IO backend.
  - Reads build plan data from a configurable Excel workbook using ExcelJS.
  - Emits `dashboardData` events to all connected clients on file change.
  - Chokidar watcher monitors the Excel file path for saves.
  - Midnight rollover timer re-emits data so the day boundary updates automatically.
  - Workstation-based grouping with priority sort.
- `public/index.html`: Single-page dashboard frontend.
  - Work in Process panel grouped by workstation.
  - Completed Today panel for finished and QC-waiting units.
  - Tailwind CSS (CDN) for layout and utility styling.
  - Day.js (CDN) for date formatting.
  - Socket.IO client for live data reception.
- `public/AZTA_BIG-7b7db612.png`: Azenta branding logo for dashboard header.
- `package.json`: Project manifest with `start` script (`node server.js`).
- `server-debug.js`: Debug variant of the server with verbose Excel-read tracing.
- `test-server.js`, `test.js`: Smoke tests for server connectivity and data reading.
- `.gitignore`: Excludes `node_modules/`, log files, temp files, and local Excel file paths.
- `CHANGELOG.md`: This file.

### Notes
- Project created to replace a static whiteboard with a live auto-refreshing display screen.
- Data source is a shared Excel workbook maintained by the production team.
- Intended to run unattended on a network-connected display (TV/monitor).
