import type { DrawingAnchor as DukeDrawingAnchor, ImageDrawing as DukeImageDrawing, Workbook } from "./duke-sheets/generated/duke_sheets_wasm.js";
import { collectWorkbookFormControls, dukeDrawingAnchorToXlsxAnchor, parseWorkbookImageAssets, type WorkbookImageAssets, type WorkbookImageSheetOrigin } from "./images";
import type { XlsxImage } from "./types";
import { createWorkbookObjectUrl } from "./workbook-object-urls";

function inferImageMimeType(source: string) {
  if (source.startsWith("data:")) {
    const separatorIndex = source.indexOf(";");
    if (separatorIndex > 5) {
      return source.slice(5, separatorIndex);
    }
  }

  const normalized = source.split("?")[0]?.toLowerCase() ?? "";
  if (normalized.endsWith(".gif")) {
    return "image/gif";
  }
  if (normalized.endsWith(".jpg") || normalized.endsWith(".jpeg")) {
    return "image/jpeg";
  }
  if (normalized.endsWith(".svg")) {
    return "image/svg+xml";
  }
  if (normalized.endsWith(".webp")) {
    return "image/webp";
  }

  return "image/png";
}

function inferWorksheetDirectImageMimeType(info: DukeImageDrawing["image"]) {
  const format = info.format;
  if (format === "gif") {
    return "image/gif";
  }
  if (format === "jpeg") {
    return "image/jpeg";
  }
  if (format === "svg") {
    return "image/svg+xml";
  }
  if (format === "bmp") {
    return "image/bmp";
  }
  if (format === "tiff") {
    return "image/tiff";
  }
  if (format === "emf") {
    return "image/emf";
  }
  if (format === "wmf") {
    return "image/wmf";
  }
  if (format === "png") {
    return "image/png";
  }

  const mediaPath = typeof info.mediaPath === "string" ? info.mediaPath : "";
  if (mediaPath) {
    return inferImageMimeType(mediaPath);
  }

  return "image/png";
}

function createWorksheetDirectImageSource(
  data: unknown,
  mimeType: string,
  objectUrls: string[]
) {
  const bytes = data instanceof Uint8Array
    ? data
    : Array.isArray(data)
      ? Uint8Array.from(data.filter((value): value is number => typeof value === "number"))
      : null;
  if (!bytes || bytes.byteLength === 0) {
    return null;
  }

  const blobBuffer = new Uint8Array(bytes.byteLength);
  blobBuffer.set(bytes);
  const objectUrl = createWorkbookObjectUrl(new Blob([blobBuffer.buffer], { type: mimeType }));
  objectUrls.push(objectUrl);
  return objectUrl;
}

function buildWorksheetDirectApiImage(
  worksheet: ReturnType<Workbook["getSheet"]>,
  workbookSheetIndex: number,
  drawing: DukeImageDrawing & { anchor: DukeDrawingAnchor },
  objectUrls: string[]
): XlsxImage | null {
  const hasSvgCompanion = Boolean(drawing.image.svgMediaPath);
  let mimeType = inferWorksheetDirectImageMimeType(drawing.image);
  let data: Uint8Array;
  try {
    const svgData = drawing.image.format === "svg" || hasSvgCompanion
      ? worksheet.drawingSvgData(drawing.drawingPath)
      : undefined;
    if (svgData) {
      data = svgData;
      mimeType = "image/svg+xml";
    } else {
      data = worksheet.drawingImageData(drawing.drawingPath);
    }
  } catch {
    return null;
  }
  const src = createWorksheetDirectImageSource(data, mimeType, objectUrls);
  if (!src) {
    return null;
  }

  const zIndex = (drawing.drawingPath[0] ?? 0) + 1;
  return {
    anchor: dukeDrawingAnchorToXlsxAnchor(drawing.anchor),
    description: drawing.altText,
    editable: false,
    id: `worksheet-image-${workbookSheetIndex}-${drawing.drawingPath.join("-")}`,
    mediaPath: drawing.image.mediaPath || undefined,
    mimeType,
    name: drawing.name,
    sheetIndex: workbookSheetIndex,
    src,
    workbookSheetIndex,
    zIndex
  };
}

function collectWorksheetApiImages(workbook: Workbook, objectUrls: string[]) {
  const directImagesByWorkbookSheetIndex = Array.from({ length: workbook.sheetCount }, () => [] as XlsxImage[]);
  for (let workbookSheetIndex = 0; workbookSheetIndex < workbook.sheetCount; workbookSheetIndex += 1) {
    const worksheet = workbook.getSheet(workbookSheetIndex);
    const rawImages = worksheet.images;
    if (rawImages.length === 0) {
      continue;
    }

    const nextImages = rawImages
      .filter((drawing): drawing is DukeImageDrawing & { anchor: DukeDrawingAnchor } => !drawing.hidden && Boolean(drawing.anchor))
      .map((drawing) => buildWorksheetDirectApiImage(worksheet, workbookSheetIndex, drawing, objectUrls))
      .filter((image): image is XlsxImage => Boolean(image));
    if (nextImages.length > 0) {
      directImagesByWorkbookSheetIndex[workbookSheetIndex] = nextImages;
    }
  }

  return directImagesByWorkbookSheetIndex;
}

function imageAnchorKey(anchor: XlsxImage["anchor"]) {
  if (anchor.kind === "absolute") {
    return [
      "absolute",
      Math.round(anchor.positionEmu.x),
      Math.round(anchor.positionEmu.y),
      Math.round(anchor.sizeEmu.cx),
      Math.round(anchor.sizeEmu.cy)
    ].join(":");
  }
  if (anchor.kind === "one-cell") {
    return [
      "one",
      anchor.from.col,
      anchor.from.row,
      Math.round(anchor.from.colOffsetEmu),
      Math.round(anchor.from.rowOffsetEmu),
      Math.round(anchor.sizeEmu.cx),
      Math.round(anchor.sizeEmu.cy)
    ].join(":");
  }
  return [
    "two",
    anchor.from.col,
    anchor.from.row,
    Math.round(anchor.from.colOffsetEmu),
    Math.round(anchor.from.rowOffsetEmu),
    anchor.to.col,
    anchor.to.row,
    Math.round(anchor.to.colOffsetEmu),
    Math.round(anchor.to.rowOffsetEmu)
  ].join(":");
}

function collectHiddenWorksheetImageAnchors(workbook: Workbook) {
  return Array.from({ length: workbook.sheetCount }, (_, workbookSheetIndex) => new Set(
    workbook.getSheet(workbookSheetIndex).images.flatMap((drawing) => (
      drawing.hidden && drawing.anchor
        ? [imageAnchorKey(dukeDrawingAnchorToXlsxAnchor(drawing.anchor))]
        : []
    ))
  ));
}

function mergeParsedAndApiImages(
  parsedImages: XlsxImage[],
  apiImages: XlsxImage[],
  hiddenAnchorKeys: Set<string>
) {
  const visibleParsedImages = parsedImages.filter((image) => !hiddenAnchorKeys.has(imageAnchorKey(image.anchor)));
  if (visibleParsedImages.length === 0) {
    return apiImages;
  }
  if (apiImages.length === 0) {
    return visibleParsedImages;
  }

  const normalizeTextKey = (value: string | undefined) => value?.trim().toLowerCase() ?? "";
  const imageKeys = (image: XlsxImage) => {
    const keys = [
      `${normalizeTextKey(image.mediaPath)}|${normalizeTextKey(image.name)}|${imageAnchorKey(image.anchor)}`,
      `${normalizeTextKey(image.mediaPath)}|${imageAnchorKey(image.anchor)}`,
      `${normalizeTextKey(image.name)}|${imageAnchorKey(image.anchor)}`,
      `${imageAnchorKey(image.anchor)}`
    ];
    return keys.filter((key, index) => key && keys.indexOf(key) === index);
  };

  const apiBuckets = new Map<string, XlsxImage[]>();
  for (const apiImage of apiImages) {
    for (const key of imageKeys(apiImage)) {
      const bucket = apiBuckets.get(key);
      if (bucket) {
        bucket.push(apiImage);
      } else {
        apiBuckets.set(key, [apiImage]);
      }
    }
  }

  const usedApiImages = new Set<XlsxImage>();
  const takeApiMatch = (image: XlsxImage) => {
    for (const key of imageKeys(image)) {
      const bucket = apiBuckets.get(key);
      if (!bucket) {
        continue;
      }
      const match = bucket.find((candidate) => !usedApiImages.has(candidate));
      if (match) {
        usedApiImages.add(match);
        return match;
      }
    }
    return null;
  };

  const merged = visibleParsedImages.map((image) => {
    const apiImage = takeApiMatch(image);
    if (!apiImage) {
      return image;
    }

    return {
      ...image,
      anchor: apiImage.anchor,
      mediaPath: apiImage.mediaPath ?? image.mediaPath,
      mimeType: apiImage.mimeType,
      name: apiImage.name ?? image.name,
      src: image.pictureRenderStatus ? image.src : apiImage.src,
      zIndex: apiImage.zIndex
    };
  });

  return merged;
}

function isZipWorkbook(bytes: Uint8Array) {
  return bytes.byteLength >= 4 && bytes[0] === 0x50 && bytes[1] === 0x4b;
}

function isLegacyXlsWorkbook(bytes: Uint8Array) {
  return bytes.byteLength >= 8
    && bytes[0] === 0xd0
    && bytes[1] === 0xcf
    && bytes[2] === 0x11
    && bytes[3] === 0xe0
    && bytes[4] === 0xa1
    && bytes[5] === 0xb1
    && bytes[6] === 0x1a
    && bytes[7] === 0xe1;
}

function shouldSkipXmlParsingForWorkbook(bytes: Uint8Array, skipXmlParsing = false) {
  return skipXmlParsing || isLegacyXlsWorkbook(bytes);
}

function createBasicWorkbookAssets(workbook: Workbook): WorkbookImageAssets {
  const objectUrls: string[] = [];
  return {
    archive: {},
    dirtyArchivePaths: new Set(),
    formControlsByWorkbookSheetIndex: collectWorkbookFormControls(workbook),
    imageOriginsById: new Map(),
    imagesByWorkbookSheetIndex: collectWorksheetApiImages(workbook, objectUrls),
    namedCellStyleByName: {},
    objectUrls,
    shapesByWorkbookSheetIndex: Array.from({ length: workbook.sheetCount }, () => []),
    sheetOrigins: Array.from({ length: workbook.sheetCount }, () => null as WorkbookImageSheetOrigin | null),
    sheetStatesByWorkbookSheetIndex: Array.from({ length: workbook.sheetCount }, () => null),
    styleById: {},
    tableMetadataByWorkbookSheetIndex: Array.from({ length: workbook.sheetCount }, () => []),
    tableStyleByName: {},
    themePalette: { colorsByIndex: {} }
  };
}

export async function loadWorkbookImageAssets(bytes: Uint8Array, workbook: Workbook, skipXmlParsing = false, archive?: Record<string, Uint8Array>) {
  if (shouldSkipXmlParsingForWorkbook(bytes, skipXmlParsing) || !isZipWorkbook(bytes)) {
    return createBasicWorkbookAssets(workbook);
  }

  const parsedAssets = await parseWorkbookImageAssets(bytes, workbook, archive);
  const apiImagesByWorkbookSheetIndex = collectWorksheetApiImages(workbook, parsedAssets.objectUrls);
  const hiddenImageAnchorsByWorkbookSheetIndex = collectHiddenWorksheetImageAnchors(workbook);

  const imagesByWorkbookSheetIndex = Array.from(
    { length: Math.max(workbook.sheetCount, parsedAssets.imagesByWorkbookSheetIndex.length, apiImagesByWorkbookSheetIndex.length) },
    (_, index) => {
      const parsedImages = parsedAssets.imagesByWorkbookSheetIndex[index] ?? [];
      const apiImages = apiImagesByWorkbookSheetIndex[index] ?? [];
      return mergeParsedAndApiImages(
        parsedImages,
        apiImages,
        hiddenImageAnchorsByWorkbookSheetIndex[index] ?? new Set()
      );
    }
  );

  return {
    ...parsedAssets,
    imagesByWorkbookSheetIndex
  };
}

