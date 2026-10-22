export type BuiltinChartColorTransform = {
  type: "tint" | "shade";
  value: number;
};

export type BuiltinChartColor = {
  themeIndex: number;
  transforms: BuiltinChartColorTransform[];
};

export type BuiltinChartSurfaceStyle = {
  floor: { fill: BuiltinChartColor | null; line: BuiltinChartColor | null };
  wall: { fill: BuiltinChartColor | null; line: BuiltinChartColor | null };
  majorGridline: BuiltinChartColor;
  minorGridline: BuiltinChartColor;
};

function decodeSrgb(channel: number) {
  return channel <= 0.04045 ? channel / 12.92 : ((channel + 0.055) / 1.055) ** 2.4;
}

function encodeSrgb(channel: number) {
  return channel <= 0.0031308 ? channel * 12.92 : 1.055 * channel ** (1 / 2.4) - 0.055;
}

function quantizeLinear(channel: number) {
  return Math.round(channel * 100000) / 100000;
}

/** DrawingML tint and shade operate on linear RGB and retain the given fraction. */
export function resolveBuiltinChartColor(
  color: BuiltinChartColor | null | undefined,
  colorsByIndex: Readonly<Record<number, string | null | undefined>>
): string | undefined {
  if (!color) return undefined;
  const hex = colorsByIndex[color.themeIndex]?.replace(/^#/, "");
  if (!hex || !/^[\da-f]{6}$/i.test(hex)) return undefined;
  if (color.transforms.length === 0) return `#${hex.toLowerCase()}`;

  const channels = [0, 2, 4].map((offset) => Math.fround(quantizeLinear(decodeSrgb(Number.parseInt(hex.slice(offset, offset + 2), 16) / 255))));
  for (const transform of color.transforms) {
    const amount = Math.fround(Math.max(0, Math.min(1, transform.value)));
    for (let index = 0; index < channels.length; index += 1) {
      const channel = channels[index]!;
      channels[index] = transform.type === "tint"
        ? Math.fround(Math.fround(1 - amount) + Math.fround(amount * channel))
        : Math.fround(amount * channel);
    }
  }
  return `#${channels.map((channel) => Math.round(encodeSrgb(quantizeLinear(channel)) * 255).toString(16).padStart(2, "0")).join("")}`;
}

const GRAYSCALE_TINTS = [88, 55, 78, 92, 70, 30];
const DARK_GRAYSCALE_TINTS = [5, 55, 78, 15, 70, 30];

function normalizeClassicStyle(styleId: number | undefined) {
  if (styleId === undefined || !Number.isInteger(styleId)) return undefined;
  if (styleId >= 1 && styleId <= 48) return styleId;
  if (styleId >= 101 && styleId <= 148) return styleId - 100;
  return undefined;
}

function themeColor(themeIndex: number, tint?: number): BuiltinChartColor {
  return {
    themeIndex,
    transforms: tint === undefined ? [] : [{ type: "tint", value: tint / 100 }]
  };
}

function appendBrightness(color: BuiltinChartColor, percent: number) {
  if (percent > 0) color.transforms.push({ type: "tint", value: (100 - percent) / 100 });
  if (percent < 0) color.transforms.push({ type: "shade", value: (100 + percent) / 100 });
  return color;
}

/** Returns defaults only; explicit colors and custom color styles take precedence. */
export function getBuiltinChartColor(
  styleId: number | undefined,
  index: number,
  count: number
): BuiltinChartColor | undefined {
  const style = normalizeClassicStyle(styleId);
  if (style === undefined || !Number.isInteger(index) || index < 0 || !Number.isInteger(count) || count < 1) {
    return undefined;
  }

  const palette = style === 41 ? 9 : ((style - 1) % 8) + 1;
  const monochrome = palette >= 3 && palette <= 8;
  const color = monochrome
    ? themeColor(palette + 1)
    : palette === 2
      ? themeColor(4 + (index % 6))
      : themeColor(1, (palette === 9 ? DARK_GRAYSCALE_TINTS : GRAYSCALE_TINTS)[index % 6]);

  const numerator = 140 * (monochrome ? index + 1 : Math.floor(index / 6) + 1);
  const denominator = monochrome ? count + 1 : Math.floor(count / 6) + 2;
  const brightness = (Math.trunc(Math.fround(Math.fround(numerator) / Math.fround(denominator))) - 70) % 100;
  return appendBrightness(color, brightness);
}

export function getBuiltinChartSurfaceStyle(styleId: number | undefined): BuiltinChartSurfaceStyle | undefined {
  const style = normalizeClassicStyle(styleId);
  if (style === undefined) return undefined;

  let fill: BuiltinChartColor | null = null;
  if (style === 33 || style === 34) fill = themeColor(1, 20);
  if (style >= 35 && style <= 40) fill = themeColor(style - 31, (styleId ?? 0) >= 100 ? 20 : undefined);
  if (style >= 41) fill = themeColor(1, 95);

  return {
    floor: { fill, line: style <= 40 ? themeColor(1, 75) : null },
    wall: { fill, line: null },
    majorGridline: themeColor(1, (styleId ?? 0) >= 100 ? 75 : undefined),
    minorGridline: themeColor(1, style >= 41 ? 90 : 50)
  };
}

export function getBuiltinChartSurfaceLine(
  styleId: number | undefined,
  index: number,
  count: number
): BuiltinChartColor | undefined {
  const style = normalizeClassicStyle(styleId);
  if (style === undefined || !Number.isInteger(index) || index < 0 || !Number.isInteger(count) || count < 1) {
    return undefined;
  }
  if (style >= 9 && style <= 16) return themeColor(2);
  if (styleId === 133) return themeColor(1, 92);
  if (style === 33 || (style >= 35 && style <= 40)) {
    return {
      themeIndex: style === 33 ? 1 : style - 31,
      transforms: [{ type: "shade", value: 0.5 }]
    };
  }
  return undefined;
}
