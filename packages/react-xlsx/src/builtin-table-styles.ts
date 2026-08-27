import type {
  XlsxResolvedCellStyle,
  XlsxTableStyleDefinition,
  XlsxThemePalette
} from "./types";

type WorkbookColor = {
  theme: number;
  tint?: number;
};

const themeColor = (theme: number, tint?: number): WorkbookColor => ({
  theme,
  ...(tint !== undefined ? { tint } : {})
});

const solidFill = (color: WorkbookColor): XlsxResolvedCellStyle["fill"] => ({
  fillType: "solid",
  color
});

type BuiltinTableStyleFamily = "Light" | "Medium" | "Dark";

type BuiltinTableStyleMatch = {
  family: BuiltinTableStyleFamily;
  index: number;
};

function parseHexColor(color: string | undefined) {
  if (!color) {
    return null;
  }

  const normalized = color.trim().replace(/^#/, "");
  if (!/^[0-9a-fA-F]{6}$/.test(normalized)) {
    return null;
  }

  const r = Number.parseInt(normalized.slice(0, 2), 16);
  const g = Number.parseInt(normalized.slice(2, 4), 16);
  const b = Number.parseInt(normalized.slice(4, 6), 16);
  return { r, g, b };
}

function channelToLinear(value: number) {
  const srgb = value / 255;
  return srgb <= 0.03928
    ? srgb / 12.92
    : ((srgb + 0.055) / 1.055) ** 2.4;
}

function resolveColorLuminance(color: string | undefined) {
  const parsed = parseHexColor(color);
  if (!parsed) {
    return null;
  }

  const r = channelToLinear(parsed.r);
  const g = channelToLinear(parsed.g);
  const b = channelToLinear(parsed.b);
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

function useWhiteBodyBanding(accentThemeIndex: number, themePalette: XlsxThemePalette) {
  const accentColor = themePalette.colorsByIndex[accentThemeIndex];
  const luminance = resolveColorLuminance(accentColor);
  if (luminance === null) {
    return false;
  }

  // Yellow-like accents are visually bright and read closer to Excel with white alternate rows.
  return luminance > 0.5;
}

function parseBuiltinTableStyleName(name: string): BuiltinTableStyleMatch | null {
  const match = /^TableStyle(Light|Medium|Dark)(\d+)$/.exec(name);
  if (!match) {
    return null;
  }

  const family = match[1] as BuiltinTableStyleFamily;
  const index = Number(match[2]);
  if (!Number.isFinite(index) || index < 1) {
    return null;
  }

  return { family, index };
}

function resolveAccentTheme(index: number) {
  // Excel built-in style galleries are arranged as 7-column blocks:
  // neutral, accent1, accent2, accent3, accent4, accent5, accent6.
  // We keep the neutral slot mapped to accent1 as a semantic fallback.
  const normalized = Math.max(index, 1);
  const slot = (normalized - 1) % 7;
  const accentOffset = slot === 0 ? 0 : slot - 1;
  return 4 + accentOffset;
}

function resolveStyleTone(index: number) {
  const normalized = Math.max(index, 1);
  return Math.floor((normalized - 1) / 7);
}

function createLightStyle(index: number): XlsxTableStyleDefinition {
  const accent = resolveAccentTheme(index);
  const tone = resolveStyleTone(index);
  const stripeTintByTone = [0.92, 0.88, 0.84, 0.8];
  const stripeTint = stripeTintByTone[tone] ?? 0.78;

  return {
    wholeTable: {},
    headerRow: {
      font: {
        bold: true,
        color: themeColor(accent)
      },
      border: {
        bottom: {
          style: "thin",
          color: themeColor(accent)
        }
      }
    },
    firstRowStripe: {
      fill: solidFill(themeColor(accent, stripeTint))
    },
    secondRowStripe: {},
    firstColumn: {
      font: {
        bold: true
      }
    },
    lastColumn: {
      font: {
        bold: true
      }
    },
    totalRow: {
      font: {
        bold: true
      },
      border: {
        top: {
          style: "thin",
          color: themeColor(accent)
        }
      }
    }
  };
}

function createMediumStyle(index: number, themePalette: XlsxThemePalette): XlsxTableStyleDefinition {
  const accent = resolveAccentTheme(index);
  const tone = resolveStyleTone(index);
  const baseTintByTone = [0.9, 0.84, 0.76, 0.68];
  const baseTint = baseTintByTone[tone] ?? 0.64;
  const stripeTintByTone = [0.8, 0.72, 0.64, 0.56];
  const stripeTint = stripeTintByTone[tone] ?? 0.52;
  const whiteBanding = useWhiteBodyBanding(accent, themePalette);

  return {
    wholeTable: whiteBanding
      ? {
          fill: solidFill(themeColor(0))
        }
      : {
          // Excel medium styles often tint the full table body, then add a stronger band.
          fill: solidFill(themeColor(accent, baseTint))
        },
    headerRow: {
      fill: solidFill(themeColor(accent)),
      font: {
        bold: true,
        color: themeColor(0)
      }
    },
    firstRowStripe: {
      fill: solidFill(themeColor(accent, stripeTint))
    },
    secondRowStripe: whiteBanding
      ? {
          fill: solidFill(themeColor(0))
        }
      : {},
    firstColumn: {
      font: {
        bold: true
      }
    },
    lastColumn: {
      font: {
        bold: true
      }
    },
    totalRow: {
      font: {
        bold: true
      },
      border: {
        top: {
          style: "double",
          color: themeColor(accent)
        }
      }
    }
  };
}

function createDarkStyle(index: number): XlsxTableStyleDefinition {
  const accent = resolveAccentTheme(index);
  const tone = resolveStyleTone(index);
  const headerDarkenByTone = [-0.2, -0.3, -0.4, -0.5];
  const headerTint = headerDarkenByTone[tone] ?? -0.35;

  return {
    wholeTable: {
      fill: solidFill(themeColor(accent, 0.75))
    },
    headerRow: {
      fill: solidFill(themeColor(accent, headerTint)),
      font: {
        bold: true,
        color: themeColor(0)
      }
    },
    firstRowStripe: {
      fill: solidFill(themeColor(accent, 0.62))
    },
    secondRowStripe: {
      fill: solidFill(themeColor(accent, 0.72))
    },
    firstColumn: {
      font: {
        bold: true
      }
    },
    lastColumn: {
      font: {
        bold: true
      }
    },
    totalRow: {
      font: {
        bold: true
      },
      border: {
        top: {
          style: "thick",
          color: themeColor(accent, headerTint)
        }
      }
    }
  };
}

function createBuiltinTableStyle(
  match: BuiltinTableStyleMatch,
  themePalette: XlsxThemePalette
): XlsxTableStyleDefinition {
  switch (match.family) {
    case "Light":
      return createLightStyle(match.index);
    case "Medium":
      return createMediumStyle(match.index, themePalette);
    case "Dark":
      return createDarkStyle(match.index);
    default:
      return createMediumStyle(match.index, themePalette);
  }
}

export function resolveBuiltinTableStyle(
  name: string | undefined,
  themePalette: XlsxThemePalette
): XlsxTableStyleDefinition | null {
  if (!name) {
    return null;
  }

  const parsed = parseBuiltinTableStyleName(name);
  if (!parsed) {
    return null;
  }

  return createBuiltinTableStyle(parsed, themePalette);
}
