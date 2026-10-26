# SharePoint version backup

Keeps **every real version** of the three shared repository files on SharePoint, so any change can be traced later
(this is how the KOULIKORO Province error was found):

| key | file | where on SharePoint |
|-----|------|---------------------|
| `admin` | `AFRO_admin_data.xlsx` | GISWORKSPACE (share link) |
| `lqas`  | `AFRO_LQAS_data_c.csv` | GISWORKSPACE (share link) |
| `im`    | `AFRO_Inside_HH_M.csv` | GISWORKSPACE / Documents / `7. SIA_Data/Data Repository` |

A new version is saved **only when the file content changed** (checked by SHA-256), so running it twice does nothing the second time.
Nothing is ever overwritten or deleted. Backups are run **manually** (double-click `run_backup.bat`).

## One-time setup

1. `pip install -r requirements.txt`
2. Copy `config/secrets.env.example` to `config/secrets.env` and paste the three `SHAREPOINT_*` values
   (the same ones used by the dashboard pipeline in the `scope` project / the IM workflow). This file is never committed.

## Everyday use

Double-click **`run_backup.bat`** (or `python sp_versions.py backup`). Run it:
- before and after anyone uploads a new admin / LQAS / IM file,
- before running a pipeline that rewrites those files.

| Command | What it does |
|---------|--------------|
| `python sp_versions.py backup [admin lqas im]` | Download the files and save a new version if they changed |
| `python sp_versions.py list [admin]` | Show saved versions (number, date, rows, size, row change, who modified) |
| `python sp_versions.py compare admin 3 5` | What changed between version 3 and 5: rows, columns, changed cells per column with first/last Excel row |
| `python sp_versions.py compare admin 3 5 --column Province --out report.xlsx` | Same, plus the top before->after changes for that column, runs of consecutive changed rows, and an Excel report listing every changed row |
| `python sp_versions.py history admin` | SharePoint's own version history (who / when). Add `--download ID` to also save one of those versions |
| `python sp_versions.py import admin <folder or file>` | Add older copies you already have (duplicates are skipped) |
| `python sp_versions.py restore admin 5 C:\temp` | Copy a saved version back out as a normal file (add `--force` to overwrite). Upload it to SharePoint yourself |

Version references can be a number, `latest`, `previous`, or the first characters of the SHA-256.

## Where things are

- `snapshots/<key>/` - the saved copies (CSV files are gzipped). **Not in git** (too large); keep this folder on a drive that is backed up.
- `catalog/versions.csv` - the index of all versions (date, rows, size, SHA-256, who modified). Kept locally next to `snapshots/` (not in git).
- `logs/` - one log per day.
- `config/tracked_files.json` - which files are tracked; add another file by adding an entry.

## Notes

- `compare` matches rows **by position** and tells you how many rows still line up on Country / District / SIA_date / Vaccine_type;
  if rows were inserted or removed the cell counts are approximate.
- Only the content on SharePoint at the moment of the backup is captured. For earlier history use `history` (SharePoint keeps older versions for a limited time).
- The IM file location assumes the `7. SIA_Data/Data Repository` folder; change `folder` in `config/tracked_files.json` if it moves.
