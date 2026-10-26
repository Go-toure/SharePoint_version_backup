#!/usr/bin/env python3
"""
SharePoint version backup - keeps every real version of the shared repository files
(Admin, LQAS, IM) so any change can be traced and compared later.

Commands (run from this folder):
  python sp_versions.py backup [admin lqas im]      download the files; save a new version only when the content changed
  python sp_versions.py list [admin]                 show the saved versions (newest last)
  python sp_versions.py compare admin 3 5 [--column Province]
                                                      what changed between two saved versions (numbers from `list`)
  python sp_versions.py history admin [--download ID] SharePoint's own version history (who changed it, when)
  python sp_versions.py import admin <folder|file>   add older copies you already have (deduplicated)
  python sp_versions.py restore admin 5 <target>     copy a saved version back out as a normal file

Nothing is ever overwritten or deleted. Snapshots go to snapshots/ (not in git); the catalog of
versions is catalog/versions.csv (small, safe to commit).
"""
import argparse, base64, csv, datetime as dt, gzip, hashlib, io, json, os, re, shutil, sys, time, zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent
CONFIG = ROOT / "config" / "tracked_files.json"
SNAP = ROOT / "snapshots"
CATALOG = ROOT / "catalog" / "versions.csv"
LOGDIR = ROOT / "logs"
FIELDS = ["version", "key", "saved_at_utc", "captured_at_utc", "sha256", "size_bytes", "rows",
          "sharepoint_modified_utc", "modified_by", "source", "snapshot", "note"]


# ----------------------------------------------------------------------------- small helpers
def log(msg):
    line = f"{dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}  {msg}"
    print(line, flush=True)
    try:
        LOGDIR.mkdir(exist_ok=True)
        with open(LOGDIR / f"backup_{dt.date.today():%Y%m%d}.log", "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except OSError:
        pass


def load_config():
    cfg = json.loads(CONFIG.read_text(encoding="utf-8"))
    return {item["key"]: item for item in cfg["files"]}


def load_secrets():
    p = ROOT / "config" / "secrets.env"
    if p.exists():
        for line in p.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, _, v = line.partition("=")
                os.environ.setdefault(k.strip(), v.strip())


def sha256_of(b):
    return hashlib.sha256(b).hexdigest()


def utcnow():
    return dt.datetime.now(dt.timezone.utc).replace(tzinfo=None)


def iso(t):
    return t.strftime("%Y-%m-%dT%H:%M:%SZ")


def parse_iso(s):
    s = (s or "").strip()
    if not s:
        return None
    try:
        return dt.datetime.fromisoformat(s.replace("Z", "+00:00")).astimezone(dt.timezone.utc).replace(tzinfo=None)
    except ValueError:
        return None


def quick_rows(data, name):
    """Row count without loading the whole file (xlsx: sheet dimension; csv: line count)."""
    try:
        if name.lower().endswith(".xlsx"):
            with zipfile.ZipFile(io.BytesIO(data)) as z:
                sheets = sorted(n for n in z.namelist() if n.startswith("xl/worksheets/sheet") and n.endswith(".xml"))
                with z.open(sheets[0]) as f:
                    head = f.read(4000).decode("utf-8", "ignore")
            m = re.search(r'<dimension ref="[A-Z]+\d+:[A-Z]+(\d+)"', head)
            return int(m.group(1)) - 1 if m else ""
        if name.lower().endswith((".csv", ".csv.gz")):
            raw = gzip.decompress(data) if name.lower().endswith(".gz") else data
            return max(raw.count(b"\n") - 1, 0)
    except Exception:
        pass
    return ""


# ----------------------------------------------------------------------------- catalog
def read_catalog():
    if not CATALOG.exists():
        return []
    with open(CATALOG, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def write_catalog(rows):
    CATALOG.parent.mkdir(exist_ok=True)
    tmp = CATALOG.with_suffix(".tmp")
    with open(tmp, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS)
        w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k, "") for k in FIELDS})
    os.replace(tmp, CATALOG)


def versions_of(rows, key):
    v = [r for r in rows if r["key"] == key]
    v.sort(key=lambda r: (r["saved_at_utc"], int(r["version"])))
    return v


def store_version(key, filename, data, saved_at, sharepoint_modified, modified_by, source, note=""):
    """Save `data` as a new version of `key` unless that exact content is already stored.
    Returns (row, is_new)."""
    rows = read_catalog()
    sha = sha256_of(data)
    for r in rows:
        if r["key"] == key and r["sha256"] == sha:
            return r, False
    n = 1 + max([int(r["version"]) for r in rows if r["key"] == key] or [0])
    stem, ext = os.path.splitext(filename)
    store_name = f"{stem}__{saved_at:%Y%m%d_%H%M%S}__{sha[:8]}{ext}"
    payload = data
    if ext.lower() == ".csv":
        store_name += ".gz"
        payload = gzip.compress(data, compresslevel=6, mtime=0)
    folder = SNAP / key
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / store_name
    tmp = path.with_suffix(path.suffix + ".part")
    tmp.write_bytes(payload)
    os.replace(tmp, path)
    row = dict(version=n, key=key, saved_at_utc=iso(saved_at), captured_at_utc=iso(utcnow()), sha256=sha,
               size_bytes=len(data), rows=quick_rows(data, filename), sharepoint_modified_utc=sharepoint_modified or "",
               modified_by=modified_by or "", source=source, snapshot=str(path.relative_to(ROOT)).replace("\\", "/"), note=note)
    rows.append(row)
    write_catalog(rows)
    return row, True


def read_snapshot_bytes(row):
    p = ROOT / row["snapshot"]
    b = p.read_bytes()
    return gzip.decompress(b) if p.name.endswith(".gz") else b


def pick_version(rows, key, ref):
    v = versions_of(rows, key)
    if not v:
        sys.exit(f"No saved versions for '{key}' yet. Run: python sp_versions.py backup {key}")
    ref = str(ref).lower()
    if ref in ("latest", "last"):
        return v[-1]
    if ref in ("previous", "prev"):
        return v[-2] if len(v) > 1 else v[-1]
    for r in v:
        if r["version"] == ref or r["sha256"].startswith(ref):
            return r
    sys.exit(f"Version '{ref}' not found for '{key}'. See: python sp_versions.py list {key}")


# ----------------------------------------------------------------------------- SharePoint (Microsoft Graph)
class Graph:
    """App-only access, same approach as the dashboard pipeline and the IM workflow upload."""
    def __init__(self):
        import requests
        self.r = requests
        t, c, s = (os.environ.get(k, "").strip() for k in ("SHAREPOINT_TENANT_ID", "SHAREPOINT_CLIENT_ID", "SHAREPOINT_CLIENT_SECRET"))
        if not all([t, c, s]):
            raise RuntimeError("SHAREPOINT_TENANT_ID / SHAREPOINT_CLIENT_ID / SHAREPOINT_CLIENT_SECRET are not set. "
                               "Copy config/secrets.env.example to config/secrets.env and fill it in "
                               "(same three values as the dashboard pipeline / IM workflow).")
        resp = requests.post(f"https://login.microsoftonline.com/{t}/oauth2/v2.0/token",
                             data={"client_id": c, "client_secret": s, "scope": "https://graph.microsoft.com/.default",
                                   "grant_type": "client_credentials"}, timeout=30)
        resp.raise_for_status()
        self.h = {"Authorization": "Bearer " + resp.json()["access_token"]}
        self._drive_cache = {}

    def get(self, url, **kw):
        last = None
        for attempt in range(3):
            try:
                resp = self.r.get(url, headers=self.h, timeout=180)
                if resp.status_code in (429, 503):
                    time.sleep(int(resp.headers.get("Retry-After", 5)))
                    continue
                resp.raise_for_status()
                return resp
            except Exception as e:  # network hiccup: retry
                last = e
                time.sleep(2 * (attempt + 1))
        raise last

    def item(self, spec):
        """driveItem metadata for a tracked file (share link or site path)."""
        if spec.get("share_url"):
            b64 = base64.b64encode(spec["share_url"].encode()).decode().rstrip("=").replace("/", "_").replace("+", "-")
            return self.get(f"https://graph.microsoft.com/v1.0/shares/u!{b64}/driveItem").json()
        site = self.get(f"https://graph.microsoft.com/v1.0/sites/{spec['hostname']}:{spec['site_path']}").json()["id"]
        if site not in self._drive_cache:
            drives = self.get(f"https://graph.microsoft.com/v1.0/sites/{site}/drives").json()["value"]
            lib = spec.get("library", "Documents").lower()
            self._drive_cache[site] = next(d["id"] for d in drives if d["name"].lower() == lib)
        drive = self._drive_cache[site]
        path = (spec["folder"].strip("/") + "/" + spec["filename"]).replace(" ", "%20")
        return self.get(f"https://graph.microsoft.com/v1.0/drives/{drive}/root:/{path}").json()

    def content(self, meta):
        url = meta.get("@microsoft.graph.downloadUrl")
        if url:
            return self.r.get(url, timeout=300).content
        return self.get(f"https://graph.microsoft.com/v1.0/drives/{meta['parentReference']['driveId']}/items/{meta['id']}/content").content

    def versions(self, meta):
        d, i = meta["parentReference"]["driveId"], meta["id"]
        return self.get(f"https://graph.microsoft.com/v1.0/drives/{d}/items/{i}/versions").json().get("value", [])

    def version_content(self, meta, vid):
        d, i = meta["parentReference"]["driveId"], meta["id"]
        return self.get(f"https://graph.microsoft.com/v1.0/drives/{d}/items/{i}/versions/{vid}/content").content


def modified_by_of(meta):
    u = (meta.get("lastModifiedBy") or {}).get("user") or {}
    return u.get("displayName") or u.get("email") or ""


# ----------------------------------------------------------------------------- commands
def cmd_backup(args, graph_factory=Graph):
    cfg = load_config()
    keys = args.keys or list(cfg)
    unknown = [k for k in keys if k not in cfg]
    if unknown:
        sys.exit(f"Unknown file key(s): {unknown}. Tracked keys: {list(cfg)}")
    load_secrets()
    try:
        g = graph_factory()
    except Exception as e:
        sys.exit(f"Could not connect to SharePoint: {e}")
    new_count = 0
    for k in keys:
        spec = cfg[k]
        try:
            meta = g.item(spec)
            data = g.content(meta)
        except Exception as e:
            log(f"[{k}] FAILED to download: {e}")
            continue
        if len(data) < int(spec.get("min_bytes", 1000)):
            log(f"[{k}] downloaded file is only {len(data)} bytes - ignored (nothing saved).")
            continue
        modified = parse_iso(meta.get("lastModifiedDateTime")) or utcnow()
        row, is_new = store_version(k, spec["filename"], data, modified, meta.get("lastModifiedDateTime", ""),
                                    modified_by_of(meta), "sharepoint")
        if is_new:
            new_count += 1
            log(f"[{k}] NEW version {row['version']}: {row['size_bytes']} bytes, {row['rows']} rows, "
                f"last modified {row['sharepoint_modified_utc']} by {row['modified_by'] or 'unknown'}")
        else:
            log(f"[{k}] unchanged (same as version {row['version']}); nothing saved.")
    log(f"Done. {new_count} new version(s).")


def human(n):
    n = float(n)
    for u in ("B", "KB", "MB", "GB"):
        if n < 1024 or u == "GB":
            return f"{n:.0f} {u}" if u == "B" else f"{n:.1f} {u}"
        n /= 1024


def cmd_list(args):
    rows = read_catalog()
    keys = args.keys or sorted({r["key"] for r in rows})
    if not rows:
        print("Nothing saved yet. Run: python sp_versions.py backup  (or: import)")
    for k in keys:
        v = versions_of(rows, k)
        print(f"\n{k}  ({len(v)} versions)")
        print(f"{'#':>3}  {'saved (UTC)':<20} {'rows':>8} {'size':>9}  {'change':<10} {'modified by':<22} source")
        prev = None
        for r in v:
            ch = ""
            if prev is not None and r["rows"] and prev["rows"]:
                d = int(r["rows"]) - int(prev["rows"])
                ch = f"{d:+d} rows" if d else "same rows"
            print(f"{r['version']:>3}  {r['saved_at_utc'].replace('T',' ').replace('Z',''):<20} {r['rows']:>8} {human(r['size_bytes']):>9}  {ch:<10} {r['modified_by'][:22]:<22} {r['source']}")
            prev = r


def _load_table(row, sheet=None):
    import pandas as pd
    data = read_snapshot_bytes(row)
    name = Path(row["snapshot"]).name
    if ".xlsx" in name.lower():
        xl = pd.ExcelFile(io.BytesIO(data))
        sh = sheet if sheet in xl.sheet_names else next((s for s in xl.sheet_names if s.lower().startswith("admin")), xl.sheet_names[0])
        df = xl.parse(sh, dtype=object)
    else:
        df = pd.read_csv(io.BytesIO(data), dtype=object, keep_default_na=True, low_memory=False)
    df = df.dropna(how="all")
    return df.reset_index(drop=True)


def cmd_compare(args):
    import pandas as pd
    rows = read_catalog()
    a, b = pick_version(rows, args.key, args.a), pick_version(rows, args.key, args.b)
    print(f"Comparing {args.key}: version {a['version']} ({a['saved_at_utc']}) -> version {b['version']} ({b['saved_at_utc']})")
    A, B = _load_table(a, args.sheet), _load_table(b, args.sheet)
    print(f"Rows: {len(A)} -> {len(B)}   Columns: {A.shape[1]} -> {B.shape[1]}")
    added, removed = [c for c in B.columns if c not in A.columns], [c for c in A.columns if c not in B.columns]
    if added:
        print("Columns added:  ", ", ".join(map(str, added)))
    if removed:
        print("Columns removed:", ", ".join(map(str, removed)))
    n = min(len(A), len(B))
    common = [c for c in A.columns if c in B.columns]
    # rows are compared by position; check that the first rows are really the same records
    keycols = [c for c in ("Country", "District", "SIA_date", "Vaccine_type") if c in common]
    if keycols:
        ka = A.loc[:n - 1, keycols].astype(str).agg("|".join, axis=1)
        kb = B.loc[:n - 1, keycols].astype(str).agg("|".join, axis=1)
        same = int((ka.values == kb.values).sum())
        print(f"Row alignment: {same} of {n} rows have the same Country/District/SIA_date/Vaccine_type at the same position"
              + ("" if same == n else "  (rows were inserted or removed - the cell comparison below is approximate)"))
    out = []
    for c in common:
        x, y = A[c].iloc[:n].astype(str).fillna(""), B[c].iloc[:n].astype(str).fillna("")
        x = x.where(A[c].iloc[:n].notna(), ""); y = y.where(B[c].iloc[:n].notna(), "")
        d = (x.values != y.values).nonzero()[0]
        if len(d):
            out.append((c, len(d), int(d[0]) + 2, int(d[-1]) + 2))
    out.sort(key=lambda t: -t[1])
    print(f"\nColumns with changed cells: {len(out)}")
    print(f"{'column':<34}{'cells changed':>14}{'first Excel row':>17}{'last Excel row':>16}")
    for c, k, f, l in out[:25]:
        print(f"{str(c):<34}{k:>14}{f:>17}{l:>16}")
    if args.column and args.column in common:
        c = args.column
        x = A[c].iloc[:n].where(A[c].iloc[:n].notna(), "(blank)").astype(str)
        y = B[c].iloc[:n].where(B[c].iloc[:n].notna(), "(blank)").astype(str)
        d = x.values != y.values
        t = pd.DataFrame({"before": x[d].values, "after": y[d].values})
        print(f"\nChanges in '{c}': {int(d.sum())}")
        if len(t):
            top = t.value_counts().head(15).reset_index(name="rows")
            print(top.to_string(index=False))
            # runs of consecutive changed rows (a big block usually means one paste / fill-down)
            idx = d.nonzero()[0]
            if len(idx):
                runs, s0, p = [], idx[0], idx[0]
                for i in idx[1:]:
                    if i != p + 1:
                        runs.append((s0, p)); s0 = i
                    p = i
                runs.append((s0, p))
                runs.sort(key=lambda r: r[0] - r[1])
                print("\nLongest runs of consecutive changed rows (Excel rows):")
                for s1, e1 in runs[:5]:
                    print(f"  rows {s1 + 2} to {e1 + 2}  ({e1 - s1 + 1} rows)")
    if args.out:
        with pd.ExcelWriter(args.out) as xw:
            pd.DataFrame(out, columns=["column", "cells_changed", "first_excel_row", "last_excel_row"]).to_excel(
                xw, sheet_name="columns", index=False)
            if args.column and args.column in common:
                c = args.column
                x = A[c].iloc[:n].where(A[c].iloc[:n].notna(), "(blank)").astype(str)
                y = B[c].iloc[:n].where(B[c].iloc[:n].notna(), "(blank)").astype(str)
                d = x.values != y.values
                det = B.loc[:n - 1, keycols].loc[d].copy() if keycols else pd.DataFrame(index=range(int(d.sum())))
                det.insert(0, "excel_row", [int(i) + 2 for i in d.nonzero()[0]])
                det["before"] = x[d].values
                det["after"] = y[d].values
                det.head(1000000).to_excel(xw, sheet_name=f"changes_{c}"[:31], index=False)
        print("Report written:", args.out)


def cmd_history(args, graph_factory=Graph):
    cfg = load_config()
    if args.key not in cfg:
        sys.exit(f"Unknown key. Tracked: {list(cfg)}")
    load_secrets()
    g = graph_factory()
    meta = g.item(cfg[args.key])
    vers = g.versions(meta)
    print(f"SharePoint version history of {meta.get('name')} ({len(vers)} versions)")
    print(f"{'id':<8}{'modified (UTC)':<24}{'modified by':<28}{'size':>12}")
    for v in vers:
        u = (v.get("lastModifiedBy") or {}).get("user") or {}
        print(f"{v.get('id',''):<8}{(v.get('lastModifiedDateTime') or '')[:19]:<24}{(u.get('displayName') or '')[:27]:<28}{v.get('size',''):>12}")
    if args.download:
        v = next((x for x in vers if x.get("id") == args.download), None)
        if not v:
            sys.exit("That version id is not in the list.")
        data = g.version_content(meta, args.download)
        u = (v.get("lastModifiedBy") or {}).get("user") or {}
        row, new = store_version(args.key, cfg[args.key]["filename"], data, parse_iso(v.get("lastModifiedDateTime")) or utcnow(),
                                 v.get("lastModifiedDateTime", ""), u.get("displayName", ""), "sharepoint-history",
                                 f"SharePoint version {args.download}")
        print(("Saved as version %s." % row["version"]) if new else f"Already stored as version {row['version']}.")


def cmd_import(args):
    cfg = load_config()
    if args.key not in cfg:
        sys.exit(f"Unknown key. Tracked: {list(cfg)}")
    src = Path(args.path)
    files = sorted(src.glob("*")) if src.is_dir() else [src]
    stem = Path(cfg[args.key]["filename"]).stem.lower()
    ext = Path(cfg[args.key]["filename"]).suffix.lower()
    files = [f for f in files if f.is_file() and f.name.lower().endswith(ext) and stem in f.name.lower() and not f.name.startswith("~$")]
    if not files:
        sys.exit("No matching files found there.")
    files.sort(key=lambda f: f.stat().st_mtime)
    new = 0
    for f in files:
        data = f.read_bytes()
        m = re.search(r"_(\d{8})_(\d{6})(?=\.|$)", f.name)   # backups named ..._YYYYMMDD_HHMMSS.ext keep their own time
        if m:
            t, how = dt.datetime.strptime(m.group(1) + m.group(2), "%Y%m%d%H%M%S"), "time taken from file name"
        else:
            t, how = dt.datetime.fromtimestamp(f.stat().st_mtime, dt.timezone.utc).replace(tzinfo=None), "time taken from file date"
        row, is_new = store_version(args.key, cfg[args.key]["filename"], data, t, "", "", "import", f"imported from {f.name} ({how})")
        new += is_new
        log(f"[{args.key}] {f.name}: " + (f"added as version {row['version']}" if is_new else f"already stored (version {row['version']})"))
    log(f"Imported {new} new version(s) out of {len(files)} file(s).")


def cmd_restore(args):
    rows = read_catalog()
    r = pick_version(rows, args.key, args.version)
    target = Path(args.target)
    if target.is_dir():
        target = target / Path(load_config()[args.key]["filename"]).name
    if target.exists() and not args.force:
        sys.exit(f"{target} already exists. Choose another name or add --force.")
    target.write_bytes(read_snapshot_bytes(r))
    print(f"Version {r['version']} ({r['saved_at_utc']}) written to {target}")


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = p.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("backup"); s.add_argument("keys", nargs="*")
    s = sub.add_parser("list"); s.add_argument("keys", nargs="*")
    s = sub.add_parser("compare"); s.add_argument("key"); s.add_argument("a"); s.add_argument("b")
    s.add_argument("--column"); s.add_argument("--sheet"); s.add_argument("--out")
    s = sub.add_parser("history"); s.add_argument("key"); s.add_argument("--download")
    s = sub.add_parser("import"); s.add_argument("key"); s.add_argument("path")
    s = sub.add_parser("restore"); s.add_argument("key"); s.add_argument("version"); s.add_argument("target"); s.add_argument("--force", action="store_true")
    a = p.parse_args()
    {"backup": cmd_backup, "list": cmd_list, "compare": cmd_compare, "history": cmd_history,
     "import": cmd_import, "restore": cmd_restore}[a.cmd](a)


if __name__ == "__main__":
    main()
