using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using System.Windows.Forms;

public static class XlsxNativeClipboard
{
    private const int MaxBytes = 32 * 1024 * 1024;
    private const string DrawingFormat = "Art::GVML ClipFormat";
    private const string ViewerFormat = "React XLSX Drawing";

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool OpenClipboard(IntPtr owner);
    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool CloseClipboard();
    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool EmptyClipboard();
    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SetClipboardData(uint format, IntPtr memory);
    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr GetClipboardData(uint format);
    [DllImport("user32.dll")]
    private static extern bool IsClipboardFormatAvailable(uint format);
    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern uint RegisterClipboardFormat(string name);
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr GlobalAlloc(uint flags, UIntPtr bytes);
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr GlobalLock(IntPtr memory);
    [DllImport("kernel32.dll")]
    private static extern bool GlobalUnlock(IntPtr memory);
    [DllImport("kernel32.dll")]
    private static extern IntPtr GlobalFree(IntPtr memory);
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern UIntPtr GlobalSize(IntPtr memory);

    private static uint Format(string name)
    {
        uint id = RegisterClipboardFormat(name);
        if (id == 0) throw new Win32Exception();
        return id;
    }

    private static void Acquire(IntPtr owner)
    {
        for (int attempt = 0; attempt < 20; attempt++)
        {
            if (OpenClipboard(owner)) return;
            Thread.Sleep(25);
        }
        throw new Win32Exception(Marshal.GetLastWin32Error(), "The clipboard is busy. Try again.");
    }

    private static byte[] Read()
    {
        Acquire(IntPtr.Zero);
        try
        {
            foreach (string name in new[] { ViewerFormat, DrawingFormat })
            {
                uint format = Format(name);
                if (!IsClipboardFormatAvailable(format)) continue;
                IntPtr memory = GetClipboardData(format);
                if (memory == IntPtr.Zero) throw new Win32Exception();
                ulong size = GlobalSize(memory).ToUInt64();
                if (size == 0 || size > MaxBytes) throw new InvalidDataException("The drawing is too large or empty.");
                IntPtr pointer = GlobalLock(memory);
                if (pointer == IntPtr.Zero) throw new Win32Exception();
                byte[] bytes = new byte[(int)size];
                try { Marshal.Copy(pointer, bytes, 0, bytes.Length); }
                finally { GlobalUnlock(memory); }
                if (name == ViewerFormat)
                {
                    int length = bytes.Length;
                    while (length > 0 && bytes[length - 1] == 0) length--;
                    Array.Resize(ref bytes, length);
                }
                return bytes;
            }
            return null;
        }
        finally { CloseClipboard(); }
    }

    private static byte[] Decode(Dictionary<string, object> payload, string key)
    {
        object value;
        if (!payload.TryGetValue(key, out value) || !(value is string)) throw new InvalidDataException("Invalid clipboard data.");
        byte[] bytes = Convert.FromBase64String((string)value);
        if (bytes.Length == 0 || bytes.Length > MaxBytes) throw new InvalidDataException("The drawing is too large or empty.");
        return bytes;
    }

    private static byte[] BitmapPreview(byte[] png)
    {
        using (var input = new MemoryStream(png))
        using (var image = Image.FromStream(input))
        {
            if ((long)image.Width * image.Height > 16000000) throw new InvalidDataException("The preview is too large.");
            using (var bitmap = new Bitmap(image.Width, image.Height, PixelFormat.Format24bppRgb))
            using (var output = new MemoryStream())
            {
                using (var graphics = Graphics.FromImage(bitmap))
                {
                    graphics.Clear(Color.White);
                    graphics.DrawImage(image, 0, 0, image.Width, image.Height);
                }
                bitmap.Save(output, ImageFormat.Bmp);
                byte[] file = output.ToArray();
                byte[] dib = new byte[file.Length - 14];
                Buffer.BlockCopy(file, 14, dib, 0, dib.Length);
                return dib;
            }
        }
    }

    private static void Write(Dictionary<string, object> payload, JavaScriptSerializer json)
    {
        byte[] drawing = Decode(payload, "officeDrawing");
        byte[] png = Decode(payload, "previewImage");
        byte[] viewer = Decode(payload, "viewerDrawing");
        if (drawing.Length < 4 || drawing[0] != 0x50 || drawing[1] != 0x4b || drawing[2] != 3 || drawing[3] != 4)
            throw new InvalidDataException("Invalid Office drawing package.");
        byte[] map = Encoding.UTF8.GetBytes(json.Serialize(new Dictionary<string, string> {
            { "application/vnd.react-xlsx.drawing+json", ViewerFormat }
        }));
        var formats = new[] { Format(DrawingFormat), Format("PNG"), 8u, Format(ViewerFormat), Format("Web Custom Format Map") };
        var representations = new[] { drawing, png, BitmapPreview(png), viewer, map };
        var handles = new IntPtr[formats.Length];
        try
        {
            for (int index = 0; index < handles.Length; index++)
            {
                handles[index] = GlobalAlloc(0x0042, new UIntPtr((uint)representations[index].Length));
                if (handles[index] == IntPtr.Zero) throw new Win32Exception();
                IntPtr pointer = GlobalLock(handles[index]);
                if (pointer == IntPtr.Zero) throw new Win32Exception();
                try { Marshal.Copy(representations[index], 0, pointer, representations[index].Length); }
                finally { GlobalUnlock(handles[index]); }
            }
            using (var owner = new Form())
            {
                Acquire(owner.Handle);
                try
                {
                    if (!EmptyClipboard()) throw new Win32Exception();
                    for (int index = 0; index < handles.Length; index++)
                    {
                        if (SetClipboardData(formats[index], handles[index]) == IntPtr.Zero) throw new Win32Exception();
                        handles[index] = IntPtr.Zero;
                    }
                }
                finally { CloseClipboard(); }
            }
        }
        finally
        {
            foreach (IntPtr handle in handles) if (handle != IntPtr.Zero) GlobalFree(handle);
        }
    }

    public static void Run()
    {
        try
        {
            string input = Console.In.ReadToEnd();
            if (input.Length > 48 * 1024 * 1024) throw new InvalidDataException("Clipboard request is too large.");
            var json = new JavaScriptSerializer { MaxJsonLength = 48 * 1024 * 1024 };
            var payload = json.Deserialize<Dictionary<string, object>>(input);
            object operation;
            if (payload == null || !payload.TryGetValue("operation", out operation)) throw new InvalidDataException("Invalid clipboard operation.");
            if ((string)operation == "read")
            {
                byte[] data = Read();
                Console.Write(json.Serialize(new Dictionary<string, object> { { "data", data == null ? null : Convert.ToBase64String(data) } }));
            }
            else if ((string)operation == "write") Write(payload, json);
            else throw new InvalidDataException("Invalid clipboard operation.");
        }
        catch (Exception error)
        {
            Console.Error.WriteLine("Native clipboard operation failed: " + error.Message);
            Environment.Exit(1);
        }
    }
}
