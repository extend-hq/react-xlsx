export function formatBinaryBytes(bytes: number) {
  if (!Number.isFinite(bytes) || bytes <= 0) return "0 B";

  const units = ["B", "KB", "MB", "GB"];
  let value = bytes;
  let unitIndex = 0;
  while (value >= 1024 && unitIndex < units.length - 1) {
    value /= 1024;
    unitIndex += 1;
  }
  return `${value >= 10 || unitIndex === 0 ? value.toFixed(0) : value.toFixed(1)} ${units[unitIndex]}`;
}

export class XlsxFileSizeLimitExceededError extends Error {
  fileSizeBytes: number;
  maxFileSizeBytes: number;

  constructor(fileSizeBytes: number, maxFileSizeBytes: number) {
    super(
      `XLSX file size ${formatBinaryBytes(fileSizeBytes)} exceeds the configured limit of ${formatBinaryBytes(maxFileSizeBytes)}.`
    );
    this.name = "XlsxFileSizeLimitExceededError";
    this.fileSizeBytes = fileSizeBytes;
    this.maxFileSizeBytes = maxFileSizeBytes;
  }
}
