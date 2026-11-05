import type { XlsxImage, XlsxImageRect } from "./types";

export function resolvePicturePaintRect(image: XlsxImage, rect: XlsxImageRect): XlsxImageRect {
  const bounds = image.paintBounds;
  if (!bounds || bounds.layoutWidthPx <= 0 || bounds.layoutHeightPx <= 0) return rect;
  const scaleX = rect.width / bounds.layoutWidthPx;
  const scaleY = rect.height / bounds.layoutHeightPx;
  return {
    left: rect.left + bounds.xPx * scaleX,
    top: rect.top + bounds.yPx * scaleY,
    width: bounds.widthPx * scaleX,
    height: bounds.heightPx * scaleY
  };
}
