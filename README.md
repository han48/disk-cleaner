# Disk Clean

A small Windows desktop utility for finding large folders and moving selected
folders to the Recycle Bin.

## Run

1. Open PowerShell in this folder.
2. Run:

   ```powershell
   powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\DiskClean.ps1
   ```

3. On first run, the app scans automatically. Later runs load the saved scan.
   Use **Force scan** to rescan all fixed drives and refresh the saved data.
4. Set the minimum folder size or change the sort order to reorganize saved
   results without rescanning. The tree expands to show matching nested
   folders; check the folders to remove and confirm.
5. Check one or more safe folders in the results tree, right-click a folder,
   and choose **Ignore checked folders** to add their absolute paths to
   `.ignores`. The folders are hidden from the current results and removed from
   the saved scan. **Ignore this folder** remains available for a single folder.

The scan runs in the background, checks fixed drives, and does not follow
junctions or symbolic links. The Windows folder is always excluded. Other
folders to exclude are configured in the `.ignores` text file beside the
script, with one entry per line. The folder containing the app is also
automatically excluded.

- A folder name (for example `Program Files`) ignores folders with that name
  anywhere on scanned drives.
- An absolute path (for example `D:\Games\Keep`) ignores that folder and its
  contents.
- Blank lines and lines starting with `#` are ignored.

The provided `.ignores` file excludes common system and application folders;
edit it to suit your system, or right-click a folder in the results tree to
add it. Its contents are loaded at each scan. If it changes after a scan
outside the app's Ignore action, cleanup is cancelled and a new scan is
required.

While scanning, the status line shows the current drive and path along with
the number of folders and files checked. The animated progress bar indicates
that the scan is active; a percentage is not shown because the total number
of folders is not known in advance. Results for scanned subfolders and their
calculated sizes are stored in `DiskClean.scan-cache.json` beside the script.
The cache is reused when the ignore rules and fixed drive list are unchanged.
Folders already covered by `.ignores` are hidden from cached results as well.
Moving folders to the Recycle Bin also runs in the background; the status line
and animated progress bar show that cleanup is active and which folder is
being processed. Successfully moved folders are removed from the current
results and the saved scan cache, so the next app launch can reuse the remaining
results.
Checking a parent selects its safe descendants too; folders containing ignored
data are shown but cannot be selected.

The scanner also limits AppData results to common temporary and regenerable
cache locations (such as Temp, crash dumps, Windows caches, and GPU shader
caches); other AppData contents are excluded. Close apps that may be using
these files before cleanup. Selected folders are moved to the Recycle Bin
where possible, not permanently deleted. Scans may take a while on large or
slow drives, and inaccessible folders are skipped.

This utility does not decide whether personal files are disposable. Always
review the path before confirming cleanup.
