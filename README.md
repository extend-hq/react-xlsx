# react-xlsx

## File size limit

`XlsxViewer` and `useXlsxViewerController` now enforce a configurable workbook file size limit before parsing.

- `maxFileSizeBytes` defaults to `25 * 1024 * 1024` (`25 MB`)
- Files above that limit render a placeholder instead of loading
- `fileTooLargeState` lets you replace that placeholder

```tsx
import { XlsxViewer } from "react-xlsx";

<XlsxViewer
  file={buffer}
  maxFileSizeBytes={50 * 1024 * 1024}
  fileTooLargeState={({ displayFileName, fileSizeBytes, maxFileSizeBytes }) => (
    <div>
      <strong>{displayFileName}</strong> is too large to open here.
      <div>
        {Math.round(fileSizeBytes / (1024 * 1024))} MB of {Math.round(maxFileSizeBytes / (1024 * 1024))} MB allowed
      </div>
    </div>
  )}
/>
```
