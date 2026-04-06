"""
fetchers/fetch_bom_rainfall.py
--------------------------------
Fetches annual rainfall totals from the Bureau of Meteorology (BOM)
using their anonymous FTP service.

Source: BOM High-Quality Monthly Rainfall dataset
FTP:    ftp://ftp.bom.gov.au/anon/home/ncc/www/change/HQmonthlyR/
Info:   ftp://ftp.bom.gov.au/anon/home/ncc/www/change/HQmonthlyR/HQmonthlyR_info.pdf

The FTP directory contains one zip per station:
  {station_padded}.zip  e.g. 043091.zip
Each zip contains a CSV with monthly totals:
  {station_padded}.csv  e.g. 043091.csv
  Columns: Station, Year, Jan, Feb, Mar, Apr, May, Jun,
                          Jul, Aug, Sep, Oct, Nov, Dec, Annual

Note: BOM's Climate Data Online website blocks programmatic HTTP access
(robots.txt disallowed, CDO requires session cookies for zip downloads).
The FTP service explicitly allows anonymous access for non-commercial use.

Station numbers are stored in towns.toml as bom_station.
Coverage: All towns with a bom_station set.

Website CSV produced:
  Environment - total rainfall.csv
"""

from __future__ import annotations

import csv
import io
import json
import sys
import zipfile
from ftplib import FTP, error_perm
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from fetchers.base import BaseFetcher
from config import CACHE_DIR

BOM_FTP_HOST = "ftp.bom.gov.au"
BOM_FTP_DIR  = "/anon/home/ncc/www/change/HQmonthlyR"
FTP_TIMEOUT  = 30


class BOMRainfallFetcher(BaseFetcher):

    SOURCE_NAME      = "bom_rainfall"
    SUPPORTED_STATES = []

    def fetch_all(self):
        # ── Connect once and fetch all stations ───────────────────────────────
        towns_with_station = [
            t for t in self.applicable_towns() if t.bom_station
        ]
        towns_without = [
            t for t in self.applicable_towns() if not t.bom_station
        ]

        for t in towns_without:
            self.log.warning(f"  [{t.name}] no bom_station in towns.toml — skipping")
            self.result.towns_skipped.append(t.name)

        if not towns_with_station:
            return

        try:
            self.log.info(f"  Connecting to {BOM_FTP_HOST}...")
            ftp = FTP(timeout=FTP_TIMEOUT)
            ftp.connect(BOM_FTP_HOST)
            ftp.login("anonymous", "boomtown-indicators@uq.edu.au")
            ftp.cwd(BOM_FTP_DIR)
            self.log.info(f"  Connected. Fetching {len(towns_with_station)} stations...")
        except Exception as exc:
            self.log.error(f"  FTP connection failed: {exc}")
            self.result.add_error("ALL", f"FTP connection failed: {exc}")
            return

        for town in towns_with_station:
            self._fetch_station(ftp, town)

        try:
            ftp.quit()
        except Exception:
            pass

    def _fetch_station(self, ftp: FTP, town):
        station     = town.bom_station
        padded      = str(station).zfill(6)
        zip_name    = f"{padded}.zip"
        cache_path  = CACHE_DIR / f"bom_rainfall_{station}.zip"

        # ── Download from FTP ─────────────────────────────────────────────────
        if not cache_path.exists() or getattr(self, '_force', False):
            try:
                buf = io.BytesIO()
                ftp.retrbinary(f"RETR {zip_name}", buf.write)
                cache_path.write_bytes(buf.getvalue())
                self.log.info(
                    f"  [{town.name}] Downloaded {zip_name} "
                    f"({cache_path.stat().st_size // 1024} KB)"
                )
            except error_perm as exc:
                self.log.warning(
                    f"  [{town.name}] Station {station} ({zip_name}) not found on FTP: {exc}"
                )
                self.result.towns_failed.append(town.name)
                return
            except Exception as exc:
                self.log.warning(f"  [{town.name}] FTP download error: {exc}")
                self.result.towns_failed.append(town.name)
                return
        else:
            self.log.info(f"  [{town.name}] Using cached {zip_name}")

        # ── Parse ─────────────────────────────────────────────────────────────
        annual = self._parse_zip(cache_path, padded, town.name)
        if not annual:
            self.result.towns_failed.append(town.name)
            return

        self._write_cache(town, station, annual)

    def _parse_zip(self, zip_path: Path, padded: str, town_name: str) -> dict:
        """
        Parse BOM HQ monthly rainfall zip.

        CSV format:
          Station, Year, Jan, Feb, Mar, Apr, May, Jun,
                         Jul, Aug, Sep, Oct, Nov, Dec, Annual

        Returns { year_int: annual_mm }
        Uses the "Annual" column directly; falls back to summing months.
        Missing months are blank or contain special values.
        """
        try:
            with zipfile.ZipFile(zip_path) as zf:
                # File is named {padded}.csv inside the zip
                csv_name = f"{padded}.csv"
                if csv_name not in zf.namelist():
                    # Try any CSV
                    csv_files = [n for n in zf.namelist() if n.endswith(".csv")]
                    if not csv_files:
                        self.log.error(
                            f"  [{town_name}] No CSV in zip: {zf.namelist()}"
                        )
                        return {}
                    csv_name = csv_files[0]

                with zf.open(csv_name) as f:
                    text = f.read().decode("utf-8", errors="replace")

        except zipfile.BadZipFile as exc:
            self.log.error(f"  [{town_name}] Bad zip: {exc}")
            return {}

        result = {}
        reader = csv.reader(text.splitlines())

        # Find header row
        header = None
        rows_iter = iter(reader)
        for row in rows_iter:
            if row and row[0].strip().lower() in ("station", "stn"):
                header = [h.strip().lower() for h in row]
                break
            # Some files start straight with data — check if col 1 looks like a year
            if row and len(row) > 1:
                try:
                    yr = int(row[1])
                    if 1800 <= yr <= 2100:
                        # No header — assume: station, year, jan..dec, annual
                        header = ["station","year","jan","feb","mar","apr","may","jun",
                                  "jul","aug","sep","oct","nov","dec","annual"]
                        # Process this row too
                        self._process_row(row, header, result)
                        break
                except (ValueError, IndexError):
                    pass

        if header is None:
            self.log.error(f"  [{town_name}] Could not find header in CSV")
            return {}

        # Find column indices
        try:
            year_col   = header.index("year")
            annual_col = next(
                (i for i, h in enumerate(header) if "annual" in h), None
            )
            month_cols = [
                header.index(m) for m in
                ["jan","feb","mar","apr","may","jun",
                 "jul","aug","sep","oct","nov","dec"]
                if m in header
            ]
        except ValueError as exc:
            self.log.error(f"  [{town_name}] Header parse error: {exc}  header={header}")
            return {}

        for row in rows_iter:
            self._process_row(row, header, result,
                              year_col, annual_col, month_cols)

        if result:
            yrs = sorted(result)
            self.log.info(
                f"  [{town_name}]: {len(result)} years, "
                f"{yrs[0]}–{yrs[-1]}, latest={result[yrs[-1]]} mm"
            )
        return result

    def _process_row(self, row, header, result,
                     year_col=1, annual_col=None, month_cols=None):
        """Parse one data row into result dict."""
        if not row or len(row) <= year_col:
            return
        try:
            year = int(row[year_col])
        except (ValueError, IndexError):
            return
        if not (1800 <= year <= 2100):
            return

        # Try annual column first
        if annual_col and annual_col < len(row):
            val = row[annual_col].strip()
            if val and val not in ("", "-9999", "99999.9"):
                try:
                    result[year] = round(float(val), 1)
                    return
                except ValueError:
                    pass

        # Fall back to summing months
        if month_cols:
            total = 0.0
            valid = 0
            for c in month_cols:
                if c < len(row):
                    v = row[c].strip()
                    if v and v not in ("", "-9999", "99999.9"):
                        try:
                            total += float(v)
                            valid += 1
                        except ValueError:
                            pass
            if valid >= 10:   # at least 10 of 12 months present
                result[year] = round(total, 1)

    def _write_cache(self, town, station: str, annual: dict):
        from config import YEAR_START, YEAR_END
        values = {
            str(yr): val for yr, val in annual.items()
            if YEAR_START <= yr <= YEAR_END
        }
        if not values:
            self.log.warning(
                f"  [{town.name}] no rainfall in range {YEAR_START}–{YEAR_END}"
            )
            self.result.towns_failed.append(town.name)
            return

        out = {
            "town":        town.name,
            "state":       town.state,
            "bom_station": station,
            "source":      "Bureau of Meteorology — High-Quality Monthly Rainfall (HQmonthlyR)",
            "source_url":  f"ftp://{BOM_FTP_HOST}{BOM_FTP_DIR}/{str(station).zfill(6)}.zip",
            "note":        "Annual total rainfall in mm from HQ monthly dataset",
            "indicators":  {"rainfall": values},
        }

        out_dir  = CACHE_DIR / "rainfall"
        out_dir.mkdir(exist_ok=True)
        out_path = out_dir / f"{town.slug}_bom_rainfall.json"
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(out, f, indent=2)

        self.result.towns_ok.append(town.name)


if __name__ == "__main__":
    from logger import get_logger
    get_logger()
    result = BOMRainfallFetcher().run()
    sys.exit(0 if result.success else 1)