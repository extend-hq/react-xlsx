export const VIEWPORT_CHANGE_EVENT = "xlsx-viewport-change";

export function visibleOverlayRect(element: HTMLElement, clipElement?: HTMLElement | null) {
  const rect = element.getBoundingClientRect();
  const win = element.ownerDocument.defaultView;
  let left = Math.max(0, rect.left);
  let top = Math.max(0, rect.top);
  let right = Math.min(win?.innerWidth ?? Infinity, rect.right);
  let bottom = Math.min(win?.innerHeight ?? Infinity, rect.bottom);
  if (clipElement) {
    const clip = clipElement.getBoundingClientRect();
    left = Math.max(left, clip.left);
    top = Math.max(top, clip.top);
    right = Math.min(right, clip.right);
    bottom = Math.min(bottom, clip.bottom);
  }
  for (let parent = element.parentElement; parent; parent = parent.parentElement) {
    const style = win?.getComputedStyle(parent);
    if (!style) continue;
    const bounds = parent.getBoundingClientRect();
    if (/auto|scroll|hidden|clip/.test(style.overflowX)) {
      left = Math.max(left, bounds.left);
      right = Math.min(right, bounds.right);
    }
    if (/auto|scroll|hidden|clip/.test(style.overflowY)) {
      top = Math.max(top, bounds.top);
      bottom = Math.min(bottom, bounds.bottom);
    }
  }
  return right > left && bottom > top ? { left, top, width: right - left, height: bottom - top } : null;
}
