# Time Keeper Phase 2

## Overview
Multi-category timesheet application built with TypeScript (no framework, DOM API only).
Phase 2 extends Phase 1 with six category tabs, a gold summary dashboard, and static sub-categories.

## Architecture
- **Two deliverable files**: `src/timesheet-phase2.ts` and `src/timesheet-phase2.css`
- TypeScript compiles to `dist/timesheet-phase2.js`; CSS copies to `dist/timesheet-phase2.css`
- Simple Express static server (`server.js`) serves `public/index.html` which loads the compiled output
- Authentication: MSAL.js (Azure AD / Entra ID) loaded from CDN
- API: Azure Function backend endpoints for accounts, timesheets, users

## Project Structure
```
src/
  timesheet-phase2.ts    # Main application TypeScript
  timesheet-phase2.css   # Complete stylesheet
dist/                    # Compiled output (gitignored)
public/
  index.html             # Host page
server.js                # Express static server (port 5000)
tsconfig.json            # TypeScript configuration
```

## Key Features
- 6 category tabs: Customer Deliverable (default), Internal, Business Development, Finance, Human Resources, Product/System
- Dynamic categories (Customer Deliverable, Business Development) load account rows from API
- Sub-categories for all non-dynamic categories fetched from `GetLookupValuesDynamic` endpoint (date-filtered per selected month)
- Gold-themed (#c6a05c gradient to #d1b06f) summary dashboard with expand/collapse per category
- Synchronized horizontal scrolling between header, body, and summary
- Stage gating: Pending/Submitted/Finalised controls editing
- Debounced auto-save (500ms) — only saves rows with hours > 0
- Phase 1 visual theme preserved (Segoe UI/Roboto, glassmorphism, blue gradient)
- Weekend column shading in both grids (summary 8%, entry 14%)
- Today highlighting (summary uses 50% lighter tint of #2b3c55, entry uses gold accent)
- Sticky Category/Total columns with opaque backgrounds and z-index: 5

## Code Organisation (TS)
- Uses `#region` / `#endregion` directives for code folding
- Single unified `TimesheetRecord` interface (merged from Phase 1 TimesheetRecord + TimesheetRowModel)
- `enrichRecord()` sets category/subCategory on API records; `createEmptyRow()` for new sub-category rows
- Interfaces: `SubCategory`, `KeyNameValuePair`, `ChoiceOptionsRequest`, `LookupValuesRequest` added
- Helper functions: `parseLookupToSubCategories`, `getCategoryNameForSubCategory`, `getSubcategoriesForCategory`
- Global arrays: `TimeSheetCategories`, `SubCategories`
- Available accounts sorted alphabetically in customer selector

## CSS Conventions
- Font weights generally use 600 (not 700) across summary dashboard and grid headers
- Summary dashboard uses `linear-gradient(135deg, #c6a05c, #d1b06f)` background
- Box shadows standardised to `0 4px 12px rgba(0, 0, 0, 0.1)` across panels
- Grid uses single `ts-grid-inner` scroll container with `ts-grid-content` wrapper (header row + body)
- Account name cells: `white-space` wrapping enabled, `max-height: 35px`
- Account button wrap uses `gap: 5px`
- Customer cell padding: `6px 6px 6px 10px`

## Recent Changes
- 2026-02-19: Replaced hardcoded `CategoryName` union type with `string` alias; replaced static `ALL_CATEGORIES` array with dynamic `getAllCategories()` function deriving from fetched `TimeSheetCategories`; removed redundant `as CategoryName` casts
- 2026-02-19: Merged TimesheetRecord and TimesheetRowModel into single TimesheetRecord interface; removed rowToRecord, recordToRowModel, sumRecordHours; added enrichRecord() method
- 2026-02-18: User manual changes — font-weight normalised to 600, gradient summary background, box-shadow standardisation, grid structure simplified back to single scroll container, account selector sorted alphabetically, #region code folding, new interfaces added, account name wrapping enabled
- 2026-02-18: Category tab order changed to: Customer Deliverable, Internal, Business Development, Finance, Human Resources, Product/System
- 2026-02-18: Summary Dashboard today highlight uses 50% lighter tint of #2b3c55
- 2026-02-18: Weekend shading bleed fixes (align-items: stretch, z-index: 5, opaque backgrounds)
- 2026-02-17: Initial Phase 2 build from Phase 1 spec

## User Preferences
- No framework — vanilla TypeScript with DOM API
- Glassmorphism design language
- Phase 1 compatibility required for API patterns and stage map
- Font weight preference: 600 over 700 for most UI elements
- Prefers subtle box-shadows: `0 4px 12px rgba(0, 0, 0, 0.1)`
