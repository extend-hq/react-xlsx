import type { XlsxThemePalette } from "./types";

function normalizeHexColor(value: string) {
  const hex = value.replace(/^#/, "");
  if (hex.length === 8) {
    return `#${hex.slice(2).toLowerCase()}`;
  }
  if (hex.length === 6) {
    return `#${hex.toLowerCase()}`;
  }
  return null;
}

function parseHexColor(color: string): [number, number, number] | null {
  const normalized = normalizeHexColor(color);
  const match = normalized ? /^#([0-9a-f]{6})$/.exec(normalized) : null;
  if (!match) {
    return null;
  }

  const hex = match[1];
  return [
    Number.parseInt(hex.slice(0, 2), 16),
    Number.parseInt(hex.slice(2, 4), 16),
    Number.parseInt(hex.slice(4, 6), 16)
  ];
}

function rgbToHsl(red: number, green: number, blue: number): [number, number, number] {
  const normalizedRed = red / 255;
  const normalizedGreen = green / 255;
  const normalizedBlue = blue / 255;
  const max = Math.max(normalizedRed, normalizedGreen, normalizedBlue);
  const min = Math.min(normalizedRed, normalizedGreen, normalizedBlue);
  const lightness = (max + min) / 2;

  if (max === min) {
    return [0, 0, lightness];
  }

  const delta = max - min;
  const saturation = lightness > 0.5 ? delta / (2 - max - min) : delta / (max + min);
  let hue = 0;

  switch (max) {
    case normalizedRed:
      hue = (normalizedGreen - normalizedBlue) / delta + (normalizedGreen < normalizedBlue ? 6 : 0);
      break;
    case normalizedGreen:
      hue = (normalizedBlue - normalizedRed) / delta + 2;
      break;
    default:
      hue = (normalizedRed - normalizedGreen) / delta + 4;
      break;
  }

  return [hue / 6, saturation, lightness];
}

function hueToRgb(p: number, q: number, t: number) {
  let nextT = t;
  if (nextT < 0) {
    nextT += 1;
  }
  if (nextT > 1) {
    nextT -= 1;
  }
  if (nextT < 1 / 6) {
    return p + (q - p) * 6 * nextT;
  }
  if (nextT < 1 / 2) {
    return q;
  }
  if (nextT < 2 / 3) {
    return p + (q - p) * (2 / 3 - nextT) * 6;
  }
  return p;
}

function hslToRgb(hue: number, saturation: number, lightness: number): [number, number, number] {
  if (saturation === 0) {
    const gray = Math.round(lightness * 255);
    return [gray, gray, gray];
  }

  const q = lightness < 0.5 ? lightness * (1 + saturation) : lightness + saturation - lightness * saturation;
  const p = 2 * lightness - q;

  return [
    Math.round(hueToRgb(p, q, hue + 1 / 3) * 255),
    Math.round(hueToRgb(p, q, hue) * 255),
    Math.round(hueToRgb(p, q, hue - 1 / 3) * 255)
  ];
}

function rgbToHex(red: number, green: number, blue: number) {
  return `#${[red, green, blue]
    .map((channel) => Math.max(0, Math.min(255, channel)).toString(16).padStart(2, "0"))
    .join("")}`;
}

function applyExcelTint(baseColor: string, tint: number) {
  const rgb = parseHexColor(baseColor);
  if (!rgb || !Number.isFinite(tint) || tint === 0) {
    return normalizeHexColor(baseColor);
  }

  const [hue, saturation, lightness] = rgbToHsl(rgb[0], rgb[1], rgb[2]);
  const nextLightness = tint < 0
    ? lightness * (1 + tint)
    : lightness * (1 - tint) + tint;
  const [nextRed, nextGreen, nextBlue] = hslToRgb(hue, saturation, Math.max(0, Math.min(1, nextLightness)));
  return rgbToHex(nextRed, nextGreen, nextBlue);
}

export function resolveWorkbookColor(
  color: Record<string, unknown> | undefined,
  themePalette?: XlsxThemePalette | null
): string | null {
  if (!color) {
    return null;
  }

  const directHex = ["hex", "rgb", "argb"]
    .map((key) => color[key])
    .find((value): value is string => typeof value === "string" && value.trim().length > 0);
  if (directHex) {
    return normalizeHexColor(directHex);
  }

  const themeValue = color.theme;
  const numericTheme = typeof themeValue === "number"
    ? themeValue
    : typeof themeValue === "string" && themeValue.trim().length > 0
      ? Number(themeValue)
      : Number.NaN;
  const themeColor = Number.isFinite(numericTheme) ? themePalette?.colorsByIndex[numericTheme] ?? null : null;
  if (!themeColor) {
    return null;
  }

  const tintValue = color.tint;
  const tint = typeof tintValue === "number"
    ? tintValue
    : typeof tintValue === "string" && tintValue.trim().length > 0
      ? Number(tintValue)
      : Number.NaN;

  return Number.isFinite(tint) ? applyExcelTint(themeColor, tint) : themeColor;
}

export function resolveWorkbookFillColor(
  fill: Record<string, unknown> | undefined,
  themePalette?: XlsxThemePalette | null
): string | null {
  if (!fill) {
    return null;
  }

  if (fill.fillType === "solid") {
    return resolveWorkbookColor(
      (fill.color as Record<string, unknown> | undefined)
      ?? (fill.foreground as Record<string, unknown> | undefined)
      ?? (fill.background as Record<string, unknown> | undefined),
      themePalette
    );
  }

  if (fill.fillType === "pattern") {
    return resolveWorkbookColor(
      (fill.foreground as Record<string, unknown> | undefined)
      ?? (fill.color as Record<string, unknown> | undefined)
      ?? (fill.background as Record<string, unknown> | undefined),
      themePalette
    );
  }

  return null;
}
