# Agents.md

## Project context
- Project: ExcelFinder (WPF, .NET 8, self-contained win-x64 publish)
- Main goal: Search Excel rows by keyword + selected JSON context, with Perforce-based edit/history/merge workflow.

## Working rules from user
- Always publish to `Build` after code changes.
- When asked, create `ExcelFinder.zip` from `Build`.
- Keep self-contained deploy (`RuntimeIdentifier=win-x64`, `SelfContained=true`) so planners can run directly.
- For each functional change, append one-line summary to `ExcelFinder.md`.

## UI and behavior to preserve
- Main window
  - Save/restore window size.
  - Save/restore JSON list pane height.
  - Folder browse remembers last folder.
  - Refresh button updates JSON list and re-runs current search.
  - Search result has CO status indicator (green=checked out, red=not checked out).
- Search result context menu
  - Editor open, History, Open folder, Perforce checkout.
  - `Who?` appears only when checked out and shows client workspace(s).
- History window
  - Context menu: Diff, Download.
  - Description must display full multi-line text from `p4 filelog -t -l`.
- Editor window
  - Save, Diff, Checkout, Checkin, Revert.
  - Show checkout ON/OFF status.
  - Checkin confirm popup: editable Description + changed points.
  - Revert confirm popup: changed points.
  - Checkin prefix input is persisted.
- Merge window
  - Show mismatched cells only, editable result value, merge to `_merged` file.
  - Preserve cell type/style behavior when applying merge.
  - Open base/compare editor at selected diff cell.

## Perforce behavior
- If checkout fails due to setup, open Perforce config popup.
- Perforce config supports Client/Host/Root/Stream apply and displays full `p4 info` output.

## Versioning
- UI version label format: `v<number>`.
- Auto-increment uses `version.counter.txt` and build metadata.

## Build and package checklist
1. `dotnet build -c Release`
2. `dotnet publish ExcelFinder.csproj -c Release -o .\\Build`
3. If requested: `Compress-Archive -Path .\\Build -DestinationPath .\\ExcelFinder.zip -Force`

## Operational note
- Dotnet/Defender file locks can break build/publish; retry sequentially after lock release.
