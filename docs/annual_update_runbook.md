# Annual update: one-page runbook

For the person who owns the indicators workbook. No programming is needed.

## Once, on a new computer

1. Install **Python 3.11 or later** from <https://www.python.org/downloads/>.
   During installation tick **"Add python.exe to PATH"**.
2. **Download the program.** Open
   <https://github.com/skennedy-clark/boomtown-indicators/releases/latest>.
   Under **Assets**, click **Source code (zip)**. No GitHub account is needed.
   (A copy of the same zip is also kept on the project's shared drive.)
3. **Unzip it** somewhere outside OneDrive, for example to `C:\boomtown-indicators`.
   Right-click the zip, choose **Extract All**, and give that folder.
4. Microsoft Excel must be installed (Windows).

The first run installs everything else by itself.

## Each year

1. **Close Excel.**
2. **Double-click `Run annual update.bat`** in the `boomtown-indicators` folder.
3. Choose **last year's finished workbook** (for example
   `Indicators Data-Charts 2026.xlsx`) in the window that opens. Notes or
   working columns added to it by hand do no harm: the update writes only into
   the year columns and charts.
4. Wait about five minutes. Excel works in the background; leave it alone.
5. The **run report** opens. Next to the workbook you chose there is now a folder
   `Indicators update <year>` containing:

   | Folder / file      | What it is                                                  |
   |--------------------|-------------------------------------------------------------|
   | `1 Data`           | every source file used, by topic, with where it came from   |
   | `2 Workbook`       | the new workbook, with the new year added and charts updated |
   | `3 Web Content`    | the files for the website, one folder per town              |
   | `Run report.txt`   | what was done, and the list of things to check              |

6. Work through **"PLEASE CHECK"** in the run report. Each line names a value that
   was not written, or was written and marked bold red, because it did not pass the
   automatic checks. Look it up at the source and type it in if it is right.
7. Do the **manual items** (below), then run the update again: every run starts
   afresh from last year's workbook, so it is safe to run as often as needed.
8. Open the new workbook and look over the charts on a few town pages.
9. Upload the `3 Web Content` folder to the website (see `docs/manual_processes.md`).

## Manual items

| What | Where it goes | Why |
|------|---------------|-----|
| SALM smoothed LGA data (CSV) from the Small Area Labour Markets page | `1 Data\Employment\salm_smoothed_lga.csv` | the website does not allow automated downloads |
| Monthly rainfall for Goondiwindi (041507) and Moranbah (034035) | `regional-indicators\manual_rainfall_data.toml` | these stations are not in the SILO dataset |

Each topic folder in `1 Data` has a "read me" with step-by-step instructions.

## If something goes wrong

- **"Cannot write the working copy … open in Excel"**: close Excel and run again.
- **A step "did not finish" and the report mentions Excel**: close Excel, open Task
  Manager, Details tab, end every `EXCEL.EXE`, and run again.
- **A data source "failed"**: the source has probably moved its file (this happens
  every year or two). The rest of the update is still usable. The full log (named
  at the end of the run report) says which source; the technical notes in
  `README.md` and the source's file in `regional-indicators/fetchers/` say where it
  is downloaded from. This is the job to hand to someone technical.

For whoever looks after the code: [`HANDOVER.md`](HANDOVER.md).
