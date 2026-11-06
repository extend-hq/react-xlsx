import argparse
import hashlib
import json
import math
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def compose(rows, title, destination):
    panel, gap, margin, card_height = 230, 12, 24, 294
    card_width = panel * 2 + gap
    canvas = Image.new("RGB", (card_width * 2 + margin * 3, 100 + math.ceil(len(rows) / 2) * card_height), "#edf1f5")
    draw = ImageDraw.Draw(canvas)
    font = ImageFont.load_default(size=18)
    small = ImageFont.load_default(size=13)
    draw.text((margin, 18), title, fill="#18212c", font=ImageFont.load_default(size=26))
    draw.text((margin, 54), "Excel 16.107.3 / React XLSX 0.20.9 + Office Art 0.1.4", fill="#4a5765", font=small)
    draw.text((margin, 74), "Native Excel reference: November 5, 2026. Viewer capture: November 6, 2026. Same pixel scale.", fill="#4a5765", font=small)
    for index, row in enumerate(rows):
        x = margin + index % 2 * (card_width + margin)
        y = 100 + index // 2 * card_height
        draw.text((x, y), row["name"], fill="#18212c", font=font)
        for side, (label, image) in enumerate([("MICROSOFT EXCEL", row["reference"]), ("REACT XLSX / RUST", row["candidate"])]):
            left = x + side * (panel + gap)
            draw.text((left, y + 27), label, fill="#4a5765", font=small)
            draw.rectangle((left, y + 48, left + panel, y + 268), fill="white", outline="#d6e0e7")
            assert image.width <= panel - 2 and image.height <= 218
            canvas.paste(image, (left + 10, y + 56))
    canvas.save(destination, optimize=True)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("reference", type=Path)
    parser.add_argument("capture", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    args.output.mkdir(parents=True, exist_ok=True)
    capture = json.loads((args.capture / "results.json").read_text())
    geometry = {}
    for line in (args.reference / "excel-geometry.tsv").read_text().strip().splitlines():
        sheet, index, name, *values = line.split("tab")
        geometry[(int(sheet) - 1, int(index))] = list(map(float, values))
    rows, records = [], []
    for sheet, data in enumerate(capture["sheets"]):
        page_path = args.reference / f"native96-{sheet + 1}.png"
        page = Image.open(page_path).convert("RGB")
        for index, item in enumerate(data["images"], 1):
            left, top, width, height, rotation = geometry[(sheet, index)]
            bounds = item.get("paintBounds") or dict(xPx=0, yPx=0, widthPx=180, heightPx=100, layoutWidthPx=180, layoutHeightPx=100)
            x = (50 + left + width / 2) * 4 / 3 - bounds["layoutWidthPx"] / 2 + bounds["xPx"]
            y = (54 + top + height / 2) * 4 / 3 - bounds["layoutHeightPx"] / 2 + bounds["yPx"]
            crop = [round(x), round(y), round(x + bounds["widthPx"]), round(y + bounds["heightPx"])]
            candidate_path = args.capture / f"{sheet}-{index}.png"
            reference = page.crop(crop)
            candidate = Image.open(candidate_path).convert("RGB")
            rows.append(dict(sheet=sheet, index=index, name=item["name"], reference=reference, candidate=candidate))
            records.append(dict(sheet=sheet, index=index, name=item["name"], nativePageSha256=digest(page_path), nativeCropPx=crop, candidateSha256=digest(candidate_path), candidateSizePx=list(candidate.size), renderStatus=item.get("pictureRenderStatus")))
    compose(rows[:12], "Excel artwork / picture transforms", args.output / "excel-officeart-transforms.png")
    compose(rows[12:], "Excel artwork / picture effects", args.output / "excel-officeart-effects.png")
    selected = {(0, 2), (0, 5), (0, 7), (0, 9), (0, 12), (1, 2), (1, 5), (1, 7), (1, 8), (1, 9), (1, 10), (1, 11)}
    compose([row for row in rows if (row["sheet"], row["index"]) in selected], "Excel artwork / native comparisons", args.output / "excel-officeart-overview.png")
    provenance = dict(viewerVersion="0.20.9", officeArtVersion="0.1.4", nativeVersion="16.107.3", nativeCaptureDate="2026-11-05", viewerCaptureDate="2026-11-06", browser=capture["browser"], devicePixelRatio=capture["devicePixelRatio"], inputSha256=digest(args.reference / "office-art-parity.xlsx"), nativePdfSha256=digest(args.reference / "excel.pdf"), nativeDpi=96, nativePageOffsetPt=[50, 54], cases=records)
    (args.output / "excel-officeart-provenance.json").write_text(json.dumps(provenance, indent=2) + "\n")
    print(f"Generated {len(rows)} Excel artwork comparisons.")


if __name__ == "__main__":
    main()
