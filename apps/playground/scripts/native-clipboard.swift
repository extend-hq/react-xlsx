import AppKit
import Foundation

struct ClipboardPayload: Decodable {
    let kind: String
    let officeDrawing: Data
    let previewImage: Data
    let viewerDrawing: Data
}

do {
    let input = FileHandle.standardInput.readDataToEndOfFile()
    guard input.count <= 48 * 1024 * 1024 else { throw ClipboardError.invalidData }
    let request = try JSONSerialization.jsonObject(with: input) as? [String: Any]
    if request?["operation"] as? String == "read" {
        let pasteboard = NSPasteboard.general
        let version = pasteboard.changeCount
        var drawing: Data?
        for type in ["org.react-xlsx.drawing", "com.microsoft.Art--GVML-ClipFormat", "com.microsoft.ShadowWorkbook"] {
            if let data = pasteboard.data(forType: NSPasteboard.PasteboardType(type)) {
                guard data.count <= 32 * 1024 * 1024 else { throw ClipboardError.invalidData }
                drawing = data
                break
            }
        }
        guard version == pasteboard.changeCount else { throw ClipboardError.changed }
        let result: [String: Any] = ["data": drawing?.base64EncodedString() as Any? ?? NSNull()]
        FileHandle.standardOutput.write(try JSONSerialization.data(withJSONObject: result))
        exit(0)
    }
    guard request?["operation"] as? String == "write" else { throw ClipboardError.invalidData }
    let payload = try JSONDecoder().decode(ClipboardPayload.self, from: input)
    guard ["chart", "image"].contains(payload.kind),
          payload.officeDrawing.starts(with: [0x50, 0x4b, 0x03, 0x04]),
          payload.previewImage.starts(with: [137, 80, 78, 71, 13, 10, 26, 10]) else {
        throw ClipboardError.invalidData
    }
    let shadow = try JSONSerialization.data(withJSONObject: [
        "shadowWorkbookBase64": payload.officeDrawing.base64EncodedString(),
        "floatingObjectId": "{\(UUID().uuidString)}",
        "floatingObjectType": payload.kind == "chart" ? 0 : 1,
    ])
    let map = try JSONSerialization.data(withJSONObject: [
        "application/vnd.react-xlsx.drawing+json": "org.react-xlsx.drawing",
        "data/shadow-workbook": "com.microsoft.ShadowWorkbook",
    ])
    let item = NSPasteboardItem()
    let representations: [(String, Data)] = [
        ("com.microsoft.Art--GVML-ClipFormat", payload.officeDrawing),
        ("public.png", payload.previewImage),
        ("org.react-xlsx.drawing", payload.viewerDrawing),
        ("com.microsoft.ShadowWorkbook", shadow),
        ("org.w3.web-custom-format.map", map),
    ]
    for (type, data) in representations {
        guard item.setData(data, forType: NSPasteboard.PasteboardType(type)) else {
            throw ClipboardError.writeFailed
        }
    }
    let pasteboard = NSPasteboard.general
    pasteboard.clearContents()
    guard pasteboard.writeObjects([item]) else { throw ClipboardError.writeFailed }
} catch {
    FileHandle.standardError.write(Data("Native clipboard operation failed: \(error)\n".utf8))
    exit(1)
}

enum ClipboardError: Error {
    case invalidData
    case writeFailed
    case changed
}
