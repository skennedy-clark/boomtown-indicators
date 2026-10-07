"""
assemble_maps/postcode_plot.py

Generates a map of a single ABS Postal Area (POA) boundary, highlighted
over a street basemap, for use as a figure in the reports of the
research project "Cumulative social and economic impacts of CSG
development in [town name]".

Input:
    ABS POA 2021 shapefile; default path
    POA_2021_AUST_GDA2020_SHP/POA_2021_AUST_GDA2020.shp, relative to
    the working directory.

Output:
    [postcode]_postcode_boundary_[basemap].png, for example
        4405_postcode_boundary_osm.png
        4405_postcode_boundary_cartodb.png

Usage (command line):
    # Default basemap (OSM) and automatic zoom
    python postcode_plot.py --postcode 4405

    # CartoDB basemap, explicit zoom level, custom output folder
    python postcode_plot.py --postcode 4405 --basemap cartodb --zoom 14 --output-dir maps/

    # List all options
    python postcode_plot.py --help

Usage (as a module):
    Import plot_postcode_boundary() and call it from another script.

    # Default basemap (OpenStreetMap): suits rural and low-density areas
    plot_postcode_boundary("4405")

    # CartoDB Voyager basemap: suits busy, built-up areas
    plot_postcode_boundary("4405", basemap="cartodb")

    # Several postcodes for one report, each with the basemap that
    # suits the town
    for poa, basemap in [("4405", "osm"), ("4350", "cartodb")]:
        plot_postcode_boundary(poa, basemap=basemap)
"""

import argparse

import geopandas as gpd
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
import contextily as cx

SHAPEFILE_PATH = "POA_2021_AUST_GDA2020_SHP/POA_2021_AUST_GDA2020.shp"

# Basemap options. Other styles are listed at:
# https://contextily.readthedocs.io/en/latest/api_reference.html#basemaps
BASEMAPS = {
    "osm": cx.providers.OpenStreetMap.Mapnik,   # more detail in rural areas; cluttered in busy, built-up areas
    "cartodb": cx.providers.CartoDB.Voyager,    # cleaner in busy, built-up areas; less detail in rural areas
}


def plot_postcode_boundary(
    postcode,
    basemap="osm",
    shapefile_path=SHAPEFILE_PATH,
    pad_fraction=0.2,
    zoom="auto",
    output_dir=".",
):
    """Plot a single POA postcode boundary over a basemap and save it as
    a PNG.

    Parameters
    ----------
    postcode : str
        The 4-digit postcode to plot, e.g. "4405". Must match the
        POA_CODE21 field in the ABS shapefile exactly (as a string).
    basemap : str, optional
        Basemap style: "osm" (default) or "cartodb".
        - "osm": OpenStreetMap Mapnik. Suits rural areas, where it shows
          more detail; can look cluttered in dense towns.
        - "cartodb": CartoDB Voyager. Cleaner, bolder labels; suits
          busy, built-up areas.
    shapefile_path : str, optional
        Path to the ABS POA shapefile (.shp).
    pad_fraction : float, optional
        Padding added on each side so that the surrounding area is
        visible, as a fraction of the boundary's longer dimension.
        0.2 = 20% padding on each side.
    zoom : int or "auto", optional
        Tile zoom level for the basemap. Higher values give more detail
        and larger labels, but can look pixelated if too high for the
        area. The default, "auto", lets contextily choose a zoom level
        from the extent being plotted. Integer values between 10 and 16
        are the useful range for a manual override.
    output_dir : str, optional
        Directory the PNG is saved into. Defaults to the current
        directory.

    Returns
    -------
    str
        The path of the saved PNG file.
    """
    if basemap not in BASEMAPS:
        raise ValueError(
            f"Unknown basemap '{basemap}'. Choose one of: {list(BASEMAPS.keys())}"
        )

    # 1. Load the ABS Postal Area (POA) shapefile.
    postcode_data = gpd.read_file(shapefile_path)

    # 2. Select the requested postcode.
    boundary = postcode_data[postcode_data["POA_CODE21"] == str(postcode)]

    if boundary.empty:
        raise SystemExit(f"Postcode {postcode} not found in the dataset.")

    # 3. Reproject to Web Mercator (EPSG:3857) so that the contextily
    # basemap tiles line up with the boundary.
    boundary_3857 = boundary.to_crs(epsg=3857)

    # 4. Compute a padded, square extent, so that the surrounding area is
    # visible and the map is square whatever the shape of the postcode.
    minx, miny, maxx, maxy = boundary_3857.total_bounds
    width = maxx - minx
    height = maxy - miny

    # The longer dimension sets the extent of both axes.
    longest_side = max(width, height)
    pad = longest_side * pad_fraction
    half_extent = (longest_side / 2) + pad

    center_x = (minx + maxx) / 2
    center_y = (miny + maxy) / 2

    xlim = (center_x - half_extent, center_x + half_extent)
    ylim = (center_y - half_extent, center_y + half_extent)

    # 5. Plot.
    fig, ax = plt.subplots(figsize=(10, 10))

    boundary_3857.plot(
        ax=ax,
        facecolor=mcolors.to_rgba("red", alpha=0.08),  # transparency applies to the fill only
        edgecolor="red",                                # fully opaque; unaffected by the fill alpha
        linewidth=1.2,
    )

    ax.set_xlim(xlim)
    ax.set_ylim(ylim)

    # Add the basemap (downloads tiles for the current view extent).
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

    # zoom is either the string "auto" or an integer zoom level.
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