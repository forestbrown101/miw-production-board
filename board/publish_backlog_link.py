"""
publish_backlog_link.py
Find the newest weekly schedules workbook (the "Backlog") in SCHEDULES\2026 and publish
its DIRECT SharePoint link for the MIW Command Center launcher.

Why: the launcher used to hunt for the file in each PC's locally synced OneDrive folder,
which fails on machines where that folder is not synced or is laid out differently. The
server always has the folder, so the server picks the newest file every rebuild and hands
out a link. Opening the link needs no local sync at all; desktop Excel opens the live,
co-authored workbook straight from SharePoint.

Ranking matches the launcher's Find-LatestWorkbook exactly: WEEK number desc, then REV
number desc, then last modified desc. Skips names containing SPENT HOURS / WORK COPY,
Excel lock files (~$) and hidden copies.

Env (all set by run_board.bat):
    BL2026         the local SCHEDULES\2026 folder
    PUB            the launcher App folder to write backlog_latest.json into
    SP_BASE        https://merchantironworks.sharepoint.com/SCHEDULES/2026/
    SP_FOLDER_URL  the SharePoint folder view, used by the launcher as a last resort
"""
import datetime, glob, json, os, re, sys, urllib.parse

FOLDER     = os.environ.get("BL2026", "")
PUB        = os.environ.get("PUB", "")
SP_BASE    = os.environ.get("SP_BASE", "https://merchantironworks.sharepoint.com/SCHEDULES/2026/")
FOLDER_URL = os.environ.get("SP_FOLDER_URL",
             "https://merchantironworks.sharepoint.com/SCHEDULES/Forms/AllItems.aspx?id=%2FSCHEDULES%2F2026")
PATTERN    = "*WEEK*SCHEDULES*.xlsx"
EXCLUDE    = ("SPENT HOURS", "WORK COPY")
OUT_NAME   = "backlog_latest.json"


def newest_workbook(folder: str):
    best = None
    for p in glob.glob(os.path.join(folder, PATTERN)):
        n = os.path.basename(p)
        if n.startswith("~$") or n.startswith("."):
            continue
        if any(x.lower() in n.lower() for x in EXCLUDE):
            continue
        stem = os.path.splitext(n)[0]
        m = re.search(r"WEEK\s*0*(\d+)", stem, re.I); wk  = int(m.group(1)) if m else 0
        m = re.search(r"REV\s*0*(\d+)",  stem, re.I); rev = int(m.group(1)) if m else 0
        key = (wk, rev, os.path.getmtime(p))
        if best is None or key > best[0]:
            best = (key, n)
    return best[1] if best else None


def atomic_write(path: str, text: str) -> None:
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        fh.write(text)
    os.replace(tmp, path)


def main() -> int:
    if not FOLDER or not os.path.isdir(FOLDER):
        print(f"backlog link: 2026 folder not found ({FOLDER!r}), leaving previous link in place")
        return 0
    if not PUB or not os.path.isdir(PUB):
        print(f"backlog link: PUB folder not found ({PUB!r})")
        return 0
    name = newest_workbook(FOLDER)
    out = os.path.join(PUB, OUT_NAME)
    if not name:
        print("backlog link: no workbook matched, leaving previous link in place")
        return 0
    url = SP_BASE + urllib.parse.quote(name)          # spaces -> %20, ampersands etc. encoded
    payload = {
        "name":       name,
        "url":        url,                             # plain link (browser / Excel Online)
        "excel":      "ms-excel:ofe|u|" + url,         # opens for edit in DESKTOP Excel
        "folder_url": FOLDER_URL,                      # last-resort fallback for the launcher
        "updated":    datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
    }
    atomic_write(out, json.dumps(payload, indent=2))
    print(f"backlog link: {name} -> {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
