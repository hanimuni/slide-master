"""Choropleth world map as a plain SVG that PowerPoint can open and ungroup (no patterns, filters or CSS classes).

Input: a world-atlas TopoJSON (countries-110m.json, Natural Earth 1:110m) and a CSV with the columns name,value.
Output: <out>.svg (one <path> per country, filled by value) and <out>.json (every outline as fractions of the picture,
where the labelled countries sit, the colour ramp and the value range). map_chart.ps1 draws the map from the .json as
native PowerPoint freeforms, one per country; the web `map` effect embeds the .svg.

    python make_choropleth.py countries-110m.json co2.csv out/co2 --center 150 --label China "United States of America" India

Standard library only. Projection: equirectangular, rotated so --center is the middle meridian (150 puts Asia-Pacific
in the middle). Countries cut by the map edge (Greenland, Antarctica with --center 150) are dropped rather than split.
"""
import argparse
import csv
import json
import sys
from pathlib import Path

ALIASES = {  # dataset name -> world-atlas name, for the usual mismatches
    "United States": "United States of America", "Democratic Republic of Congo": "Dem. Rep. Congo", "Congo": "Congo",
    "Central African Republic": "Central African Rep.", "South Sudan": "S. Sudan", "Dominican Republic": "Dominican Rep.",
    "Equatorial Guinea": "Eq. Guinea", "Bosnia and Herzegovina": "Bosnia and Herz.", "Czechia": "Czechia", "Cote d'Ivoire": "Côte d'Ivoire",
    "Eswatini": "eSwatini", "North Macedonia": "Macedonia", "Solomon Islands": "Solomon Is.", "Western Sahara": "W. Sahara",
    "East Timor": "Timor-Leste", "Falkland Islands": "Falkland Is.", "French Southern Territories": "Fr. S. Antarctic Lands",
}


def decode(topo):
    """TopoJSON -> {name: [ring, ...]} with rings as lists of (lon, lat)."""
    sx, sy = topo["transform"]["scale"]; tx, ty = topo["transform"]["translate"]
    arcs = []
    for arc in topo["arcs"]:
        x = y = 0; pts = []
        for dx, dy in arc:
            x += dx; y += dy; pts.append((x * sx + tx, y * sy + ty))
        arcs.append(pts)

    def ring(ids):
        out = []
        for i in ids:
            seg = arcs[i] if i >= 0 else arcs[~i][::-1]
            out += seg if not out else seg[1:]
        return out

    shapes = {}
    for g in topo["objects"]["countries"]["geometries"]:
        polys = [g["arcs"]] if g["type"] == "Polygon" else g["arcs"]
        shapes[g["properties"]["name"]] = [ring(r) for poly in polys for r in poly]
    return shapes


def ramp(colors, t):
    t = max(0.0, min(1.0, t)) * (len(colors) - 1); i = min(int(t), len(colors) - 2); f = t - i
    a, b = (tuple(int(c[k:k + 2], 16) for k in (1, 3, 5)) for c in (colors[i], colors[i + 1]))
    return "#%02X%02X%02X" % tuple(round(a[k] + (b[k] - a[k]) * f) for k in range(3))


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("topo"); ap.add_argument("data"); ap.add_argument("out", help="output path without extension")
    ap.add_argument("--center", type=float, default=150, help="middle meridian in degrees east")
    ap.add_argument("--colors", nargs="+", default=["#FFF3B0", "#FDBE5A", "#F0702E", "#C8261B", "#6E1210"], help="ramp, low to high")
    ap.add_argument("--nodata", default="#D9D9D9"); ap.add_argument("--max", type=float, default=0, help="value at the top of the ramp (default: the largest value)")
    ap.add_argument("--gamma", type=float, default=0.5, help="<1 spreads the low values apart (one giant value would flatten the rest)")
    ap.add_argument("--lat", type=float, nargs=2, default=(-56, 78), help="south and north limit")
    ap.add_argument("--only", nargs="*", default=[], help="draw just these countries, enlarged to fill the frame")
    ap.add_argument("--label", nargs="*", default=[], help="countries whose position goes into the sidecar file")
    ap.add_argument("--width", type=int, default=1600)
    a = ap.parse_args()

    shapes = decode(json.loads(Path(a.topo).read_text(encoding="utf-8")))
    values = {}
    with open(a.data, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            try:
                values[ALIASES.get(row["name"], row["name"])] = float(row["value"])
            except ValueError:
                pass
    top = a.max or max(values.values())
    south, north = a.lat
    rel = lambda lon: (lon - a.center + 180) % 360               # 0..360 across the map
    drawn = {}
    for name, rings in shapes.items():
        if a.only and name not in a.only:
            continue
        keep = []
        for r in rings:
            xs = [rel(lon) for lon, _ in r]
            if max(abs(xs[i] - xs[i - 1]) for i in range(1, len(xs))) > 180:   # cut by the map edge
                continue
            keep.append([(x, max(south, min(north, lat))) for x, (_, lat) in zip(xs, r)])
        if keep:
            drawn[name] = keep
    if not drawn:
        sys.exit("nothing to draw")
    if a.only:                                                    # frame the selection
        allp = [p for rs in drawn.values() for r in rs for p in r]
        x0, x1 = min(p[0] for p in allp), max(p[0] for p in allp); y0, y1 = min(p[1] for p in allp), max(p[1] for p in allp)
    else:
        x0, x1, y0, y1 = 0, 360, south, north
    k = a.width / (x1 - x0); height = round((y1 - y0) * k)
    P = lambda x, y: ((x - x0) * k, (y1 - y) * k)

    paths, labels, outlines = [], {}, []
    for name, rings in sorted(drawn.items()):
        v = values.get(name)
        fill = a.nodata if v is None else ramp(a.colors, (v / top) ** a.gamma)
        d = "".join("M" + "L".join("%.1f %.1f" % P(x, y) for x, y in r) + "Z" for r in rings)
        paths.append(f'<path id="{name.replace(chr(34), "")}" d="{d}" fill="{fill}" stroke="#FFFFFF" stroke-width="0.6"/>')
        # the same outlines as fractions of the picture, for drawing the map as native shapes (islands under 6 points are left out)
        outlines.append({"name": name, "fill": fill, "rings": [[[round(P(x, y)[0] / a.width, 4), round(P(x, y)[1] / height, 4)] for x, y in r] for r in rings if len(r) >= 6 or len(rings) == 1]})
        if name in a.label:
            big = max(rings, key=len)                             # label on the largest piece: average of its outline
            cx, cy = P(sum(p[0] for p in big) / len(big), sum(p[1] for p in big) / len(big))
            labels[name] = {"x": round(cx / a.width, 4), "y": round(cy / height, 4), "value": v, "dark": v is not None and (v / top) ** a.gamma > 0.55}
    svg = (f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {a.width} {height}" width="{a.width}" height="{height}">' + "".join(paths) + "</svg>")
    out = Path(a.out); out.parent.mkdir(parents=True, exist_ok=True)
    out.with_suffix(".svg").write_text(svg, encoding="utf-8")
    out.with_suffix(".json").write_text(json.dumps({"width": a.width, "height": height, "max": top, "colors": a.colors, "nodata": a.nodata, "labels": labels, "shapes": outlines,
                                                    "missing": sorted(n for n in drawn if n not in values)}, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"{len(drawn)} countries, {len([n for n in drawn if n in values])} with data -> {out.with_suffix('.svg')} ({len(svg) // 1024} KB)")


if __name__ == "__main__":
    main()
