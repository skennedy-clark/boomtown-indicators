# Postcode Boundary Map Generator

Generates a map of a single ABS Postal Area (POA) boundary, highlighted
over a street basemap. Built for producing figures in the Research
Project: *Cumulative social and economic impacts of CSG development in
[town name]* reports.

Each map shows the target postcode boundary (filled red, semi-transparent,
with a solid black/red outline) over a basemap of the surrounding area, so
readers can see the postcode in geographic context rather than as an
isolated shape.

---

## 1. Getting the boundary data

This script needs the ABS **Postal Areas — 2021 — Shapefile**, which is included in this folder, or can be downloaded separately.

1. Go to the ABS Digital Boundary Files page: https://www.abs.gov.au/statistics/standards/australian-statistical-geography-standard-asgs/edition-3-july-2021-june-2026/access-and-downloads/digital-boundary-files
2. Scroll down to the **Data files** section and find **Postal Areas — 2021 — Shapefile**.
3. Click the **Download zip** link to retrieve `POA_2021_AUST_GDA2020_SHP.zip`.
4. Unzip it so the `.shp` (and its companion `.dbf`, `.shx`, `.prj`, etc.)
   end up at:

   ```
   POA_2021_AUST_GDA2020_SHP/POA_2021_AUST_GDA2020.shp
   ```

   relative to wherever you run the script from (or update
   `SHAPEFILE_PATH` in `postcode_plot.py` to point elsewhere).

The file is published in both GDA2020 and GDA94 datums, with GDA2020 being the current official national datum — make sure you grab the **GDA2020** version specifically, as that's what this script assumes.

Ocasionally the Postal Areas will be updated and this script will need to be modified for the new data. 

---

## 2. Installing dependencies

```bash
pip install geopandas matplotlib contextily
```

`contextily` downloads basemap tiles over the internet at run time, so an
active internet connection is required when generating maps (the
shapefile itself works fully offline once downloaded).

---

## 3. Usage

### Command line (recommended)

```bash
# Default basemap (OSM) and auto zoom
python postcode_plot.py --postcode 4405

# CartoDB basemap, explicit zoom level, custom output folder
python postcode_plot.py --postcode 4405 --basemap cartodb --zoom 14 --output-dir maps/

# See all options
python postcode_plot.py --help
```

This saves a file named:

```
[postcode]_postcode_boundary_[basemap].png
```

e.g.

```
4405_postcode_boundary_osm.png
4405_postcode_boundary_cartodb.png
```

### As a Python module

Import and call the function directly:

```python
from postcode_plot import plot_postcode_boundary

# Default basemap (OpenStreetMap) — best for rural / low-density areas
plot_postcode_boundary("4405")

# CartoDB Voyager basemap — better for busy / built-up areas
plot_postcode_boundary("4405", basemap="cartodb")
```

### Batch-generating maps for a report

```python
from postcode_plot import plot_postcode_boundary

towns = [
    ("4405", "osm"),       # Dalby — rural, OSM gives more detail
    ("4350", "cartodb"),   # Toowoomba — built-up, CartoDB is cleaner
]

for postcode, basemap in towns:
    plot_postcode_boundary(postcode, basemap=basemap)
```

---

## 4. Parameters

| CLI flag             | Python parameter   | Default                                          | Description |
|------------------------|---------------------|---------------------------------------------------|--------------|
| `--postcode`           | `postcode`           | *required*                                        | 4-digit postcode string, e.g. `"4405"`. Must match `POA_CODE21` in the shapefile exactly. |
| `--basemap`            | `basemap`            | `"osm"`                                           | `"osm"` (OpenStreetMap Mapnik) or `"cartodb"` (CartoDB Voyager). |
| `--shapefile-path`     | `shapefile_path`     | `POA_2021_AUST_GDA2020_SHP/POA_2021_AUST_GDA2020.shp` | Path to the ABS shapefile. |
| `--pad-fraction`       | `pad_fraction`       | `0.2`                                             | Padding around the boundary, as a fraction of its width/height, so surrounding context is visible. |
| `--zoom`               | `zoom`               | `"auto"`                                          | Basemap tile zoom level. `"auto"` lets contextily pick an appropriate level for the extent. Pass an integer (e.g. `14`) to override manually — higher = more detail/larger labels, but can look pixelated if pushed too high. Try 10–16. |
| `--output-dir`         | `output_dir`         | `"."`                                             | Folder to save the PNG into. |

---

## 5. Choosing a basemap

- **`osm`** (OpenStreetMap Mapnik) — more detail in rural areas; can look
  cluttered in busy, built-up towns.
- **`cartodb`** (CartoDB Voyager) — cleaner, bolder labels; better suited
  to busy/built-up areas, but shows less detail in sparse rural areas.

More basemap styles are available from `contextily.providers` — see the
full list at:
<https://contextily.readthedocs.io/en/latest/api_reference.html#basemaps>

---

## 6. Notes

- The shapefile's `POA_CODE21` field is a string, so postcodes should be
  passed as strings (`"4405"`, not `4405`) to avoid type-mismatch issues
  or comparison errors.
- Basemap tiles are cached locally after first download, so re-running
  for the same area/zoom is faster on subsequent runs.
- Boundary fill uses a transparent red fill with an opaque outline (fill
  alpha and outline are set independently), so the underlying basemap
  stays legible through the highlighted area.