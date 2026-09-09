"""
postcode_plot.py

Generates a map of a single ABS Postal Area (POA) boundary, highlighted
over a street basemap, for use as figures in the Research Project:
"Cumulative social and economic impacts of CSG development in [town name]"
reports.

Each output image is saved as:
    [postcode]_postcode_boundary_[basemap].png
e.g. 4405_postcode_boundary_osm.png
     4405_postcode_boundary_cartodb.png

------------------------------------------------------------------------
USAGE - COMMAND LINE
------------------------------------------------------------------------
# Default basemap (OSM) and auto zoom
python postcode_plot.py --postcode 4405

# CartoDB basemap, explicit zoom level, custom output folder
python postcode_plot.py --postcode 4405 --basemap cartodb --zoom 14 --output-dir maps/

# Run --help to see all options
python postcode_plot.py --help

------------------------------------------------------------------------
USAGE - AS A MODULE
------------------------------------------------------------------------
Import and call plot_postcode_boundary() from another script.

Examples
--------
# Default basemap (OpenStreetMap) - best for rural/low-density areas
plot_postcode_boundary("4405")

# CartoDB Voyager basemap - better for busy/built-up areas
plot_postcode_boundary("4405", basemap="cartodb")

# Process several postcodes for a report in one go, each with the
# basemap best suited to that town
for poa, basemap in [("4405", "osm"), ("4350", "cartodb")]:
    plot_postcode_boundary(poa, basemap=basemap)
------------------------------------------------------------------------
"""

import argparse

import geopandas as gpd
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
import contextily as cx

SHAPEFILE_PATH = "POA_2021_AUST_GDA2020_SHP/POA_2021_AUST_GDA2020.shp"

# Basemap options - experiment with more styles at:
# https://contextily.readthedocs.io/en/latest/api_reference.html#basemaps
BASEMAPS = {
    "osm": cx.providers.OpenStreetMap.Mapnik,   # more detail in rural areas, busier/built-up areas get cluttered
    "cartodb": cx.providers.CartoDB.Voyager,    # cleaner for busy/built-up areas, less detail in rural areas
}


def plot_postcode_boundary(
    postcode,
    basemap="osm",
    shapefile_path=SHAPEFILE_PATH,
    pad_fraction=0.2,
    zoom="auto",
    output_dir=".",
):
    """
    Plot a single POA postcode boundary over a basemap and save to PNG.

    Parameters
    ----------
    postcode : str
        The 4-digit postcode to plot, e.g. "4405". Must match the
        POA_CODE21 field in the ABS shapefile exactly (as a string).
    basemap : str, optional
        Which basemap style to use. One of "osm" (default) or "cartodb".
        - "osm": OpenStreetMap Mapnik - best for rural areas (more detail
          where there's not much going on, can get busy in dense towns)
        - "cartodb": CartoDB Voyager - cleaner/bolder labels, better for
          busy/built-up areas
    shapefile_path : str, optional
        Path to the ABS POA shapefile (.shp).
    pad_fraction : float, optional
        Fraction of the boundary's width/height to pad on each side, so
        surrounding context is visible. 0.2 = 20% padding each side.
    zoom : int or "auto", optional
        Tile zoom level for the basemap. Higher = more detail/larger
        labels, but can look pixelated if pushed too high for the area.
        Defaults to "auto", which lets contextily pick an appropriate
        zoom level based on the extent being plotted. Try integer values
        between 10-16 if you want to override it manually.
    output_dir : str, optional
        Directory to save the output PNG into. Defaults to the current
        directory.

    Returns
    -------
    str
        The path to the saved PNG file.
    """
    if basemap not in BASEMAPS:
        raise ValueError(
            f"Unknown basemap '{basemap}'. Choose one of: {list(BASEMAPS.keys())}"
        )

    # 1. Load the ABS Postal Area (POA) shapefile
    postcode_data = gpd.read_file(shapefile_path)

    # 2. Filter for the postcode of interest
    boundary = postcode_data[postcode_data["POA_CODE21"] == str(postcode)]

    if boundary.empty:
        raise SystemExit(f"Postcode {postcode} not found in the dataset.")

    # 3. Reproject to Web Mercator (EPSG:3857) so contextily basemaps line up
    boundary_3857 = boundary.to_crs(epsg=3857)

    # 4. Work out a padded, SQUARE extent so we can see the surrounding
    # area too, and the map always comes out as a square regardless of
    # whether the postcode itself is wide, tall, or oddly shaped.
    minx, miny, maxx, maxy = boundary_3857.total_bounds
    width = maxx - minx
    height = maxy - miny

    # Use whichever dimension is longer as the basis for both axes
    longest_side = max(width, height)
    pad = longest_side * pad_fraction
    half_extent = (longest_side / 2) + pad

    center_x = (minx + maxx) / 2
    center_y = (miny + maxy) / 2

    xlim = (center_x - half_extent, center_x + half_extent)
    ylim = (center_y - half_extent, center_y + half_extent)

    # 5. Plot
    fig, ax = plt.subplots(figsize=(10, 10))

    boundary_3857.plot(
        ax=ax,
        facecolor=mcolors.to_rgba("red", alpha=0.08),  # transparency baked into fill only
        edgecolor="red",                                # fully opaque, unaffected by fill alpha
        linewidth=1.2,
    )

    ax.set_xlim(xlim)
    ax.set_ylim(ylim)

    # Add the basemap (downloads tiles for the current view extent)
    cx.add_basemap(ax, source=BASEMAPS[basemap], zoom=zoom)

    ax.set_axis_off()
    plt.tight_layout()

    output_path = f"{output_dir.rstrip('/')}/{postcode}_postcode_boundary_{basemap}.png"
    plt.savefig(output_path, dpi=150, bbox_inches="tight")
    plt.close(fig)

    print(f"Saved {output_path}")
    return output_path


def _parse_args():
    parser = argparse.ArgumentParser(
        description=(
            "Plot an ABS Postal Area (POA) boundary over a basemap and "
            "save it as a PNG."
        )
    )
    parser.add_argument(
        "--postcode",
        required=True,
        help='4-digit postcode to plot, e.g. "4405".',
    )
    parser.add_argument(
        "--basemap",
        choices=sorted(BASEMAPS.keys()),
        default="osm",
        help='Basemap style to use. Default: "osm".',
    )
    parser.add_argument(
        "--shapefile-path",
        default=SHAPEFILE_PATH,
        help=f"Path to the ABS POA shapefile (.shp). Default: {SHAPEFILE_PATH}",
    )
    parser.add_argument(
        "--pad-fraction",
        type=float,
        default=0.2,
        help="Padding around the boundary, as a fraction of its width/height. Default: 0.2",
    )
    parser.add_argument(
        "--zoom",
        default="auto",
        help='Basemap tile zoom level (integer), or "auto" to let contextily choose. Default: "auto".',
    )
    parser.add_argument(
        "--output-dir",
        default=".",
        help='Folder to save the PNG into. Default: current directory (".").',
    )
    return parser.parse_args()


def main():
    args = _parse_args()

    # zoom can be the string "auto" or an integer zoom level
    zoom = args.zoom
    if zoom != "auto":
        zoom = int(zoom)

    plot_postcode_boundary(
        postcode=args.postcode,
        basemap=args.basemap,
        shapefile_path=args.shapefile_path,
        pad_fraction=args.pad_fraction,
        zoom=zoom,
        output_dir=args.output_dir,
    )


if __name__ == "__main__":
    main()