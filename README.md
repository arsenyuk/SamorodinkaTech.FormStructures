# SamorodinkaTech.FormStructures

ASP.NET Core Razor Pages app for:
- uploading Excel templates (empty header-only files) and extracting a versioned schema
- uploading filled Excel files, extracting rows, and storing data uploads

## What it does
- **Schema upload (template)**
  - Reads:
    - Row 1: form number (first non-empty cell)
    - Row 2: form title (first non-empty cell)
    - Row 3+: header (simple or merged multi-row)
  - Builds:
    - hierarchical header tree
    - flat leaf column list (with stable `Path`)
  - Stores the original template and extracted `structure.json`
  - Versions schemas by structure hash (`StructureHash`)

- **Data upload (filled file)**
  - Parses the file schema and matches it to a stored schema version by `StructureHash`
  - Extracts non-empty data rows into `data.json`
  - Stores the original filled file (`original.xlsx`) and `meta.json`

## UI routes
- `GET /` — list of forms
- `GET /settings/forms` — upload/analyze empty templates + list of forms
- `GET /forms/{formNumber}` — schema details (versions, header tree, columns) + upload filled file + upload new schema version
- `GET /forms/{formNumber}/data` — list of all data uploads for the form
- `GET /forms/{formNumber}/data/aggregated` — aggregated data view across uploads (for latest schema version by default)
- `GET /forms/{formNumber}/data/v{version:int}/{uploadId}` — extracted data rows for one upload
- `GET /forms/{formNumber}/data/v{version:int}/{uploadId}/file` — upload file page (metadata + download file/JSON)

## Form index vs form number
The app uses two identifiers:

- **Form index** — stable internal ID used in URLs and storage folder names.
  - Example: `GET /forms/001` (index = `001`)
- **Form number** — business identifier from the Excel template (Row 1). It can be complex and does not need to look like the index.

When uploading a new schema version or a data file from `GET /forms/{index}`:
- the upload is always stored under that `{index}`
- the Excel form number from Row 1 is preserved inside the stored schema version (read-only snapshot)

You can change the displayed form title/number (metadata) on the form page without changing the form index/URL.

## Uploading new schema versions
- Upload an empty template (.xlsx) with schema changes.
- Use either:
  - `GET /settings/forms` → "Upload empty template"
  - `GET /forms/{index}` → "Upload new schema version"

If the schema hash did not change, a new version is not created.

## Storage
By default, files are stored under `SamorodinkaTech.FormStructures.Web/storage/`.

- Schemas: `SamorodinkaTech.FormStructures.Web/storage/forms/{formNumber}/v{n}/`
  - `original.xlsx`
  - `structure.json`
- Data uploads: `SamorodinkaTech.FormStructures.Web/storage/data/{formNumber}/v{n}/{uploadId}/`
  - `original.xlsx`
  - `data.json`
  - `meta.json`

## Run
From repo root:

```bash
ASPNETCORE_ENVIRONMENT=Development DOTNET_ENVIRONMENT=Development \
  dotnet watch --project SamorodinkaTech.FormStructures.Web/SamorodinkaTech.FormStructures.Web.csproj run
```

If you don't want file watching / auto-restart, use:

```bash
ASPNETCORE_ENVIRONMENT=Development DOTNET_ENVIRONMENT=Development \
  dotnet run --project SamorodinkaTech.FormStructures.Web/SamorodinkaTech.FormStructures.Web.csproj
```

Logs are written to `SamorodinkaTech.FormStructures.Web/logs/`.

Note: for proper Static Web Assets (CSS/Bootstrap), prefer running from sources in `Development` (as above).

## Tools
There is a small seeding utility that can generate demo Excel files and (optionally) store a data upload:

```bash
dotnet run --project tools/SeedForm/SeedForm.csproj -c Debug
dotnet run --project tools/SeedForm/SeedForm.csproj -c Debug -- data
```

Outputs are written under `tools/SeedForm/out/` (ignored by git).
