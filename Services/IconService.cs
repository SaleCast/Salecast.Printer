using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using SkiaSharp;
using Svg.Skia;

namespace SaleCast.Printer.Services;

/// <summary>
/// Service for loading and converting icons from various formats (SVG, PNG, ICO)
/// </summary>
public static class IconService
{
    /// <summary>
    /// Load an icon from the Assets folder, trying SVG first, then ICO, then generating one
    /// </summary>
    [SupportedOSPlatform("windows")]
    public static Icon LoadAppIcon(int size = 32)
    {
        var assetsPath = Path.Combine(AppContext.BaseDirectory, "Assets");

        // Try SVG first
        var svgPath = Path.Combine(assetsPath, "icon.svg");
        if (File.Exists(svgPath))
        {
            var icon = LoadIconFromSvg(svgPath, size);
            if (icon != null) return icon;
        }

        // Try ICO
        var icoPath = Path.Combine(assetsPath, "icon.ico");
        if (File.Exists(icoPath))
        {
            return new Icon(icoPath, size, size);
        }

        // Try PNG
        var pngPath = Path.Combine(assetsPath, "icon.png");
        if (File.Exists(pngPath))
        {
            var icon = LoadIconFromPng(pngPath, size);
            if (icon != null) return icon;
        }

        // Try to extract from exe
        var exePath = Environment.ProcessPath;
        if (exePath != null && File.Exists(exePath))
        {
            var icon = Icon.ExtractAssociatedIcon(exePath);
            if (icon != null) return icon;
        }

        // Generate fallback icon
        return CreateFallbackIcon(size);
    }

    /// <summary>
    /// Convert SVG file to Windows Icon
    /// </summary>
    [SupportedOSPlatform("windows")]
    public static Icon? LoadIconFromSvg(string svgPath, int size = 32)
    {
        try
        {
            using var svg = new SKSvg();
            svg.Load(svgPath);

            if (svg.Picture == null) return null;

            // Calculate scale to fit the desired size
            var bounds = svg.Picture.CullRect;
            var scale = Math.Min(size / bounds.Width, size / bounds.Height);

            using var bitmap = new SKBitmap(size, size);
            using var canvas = new SKCanvas(bitmap);

            canvas.Clear(SKColors.Transparent);
            canvas.Scale(scale);
            canvas.DrawPicture(svg.Picture);

            return SkBitmapToIcon(bitmap);
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Convert PNG file to Windows Icon
    /// </summary>
    [SupportedOSPlatform("windows")]
    public static Icon? LoadIconFromPng(string pngPath, int size = 32)
    {
        try
        {
            using var originalBitmap = new Bitmap(pngPath);
            using var resizedBitmap = new Bitmap(originalBitmap, new Size(size, size));
            return Icon.FromHandle(resizedBitmap.GetHicon());
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Convert SkiaSharp bitmap to Windows Icon
    /// </summary>
    [SupportedOSPlatform("windows")]
    private static Icon SkBitmapToIcon(SKBitmap skBitmap)
    {
        using var bitmap = new Bitmap(skBitmap.Width, skBitmap.Height, PixelFormat.Format32bppArgb);

        var bitmapData = bitmap.LockBits(
            new Rectangle(0, 0, skBitmap.Width, skBitmap.Height),
            ImageLockMode.WriteOnly,
            PixelFormat.Format32bppArgb);

        // Copy pixels from SkiaSharp to System.Drawing
        var pixels = skBitmap.GetPixelSpan();
        Marshal.Copy(pixels.ToArray(), 0, bitmapData.Scan0, pixels.Length);

        bitmap.UnlockBits(bitmapData);

        return Icon.FromHandle(bitmap.GetHicon());
    }

    /// <summary>
    /// Create a simple fallback icon (SaleCast-style gradient with chart)
    /// </summary>
    [SupportedOSPlatform("windows")]
    private static Icon CreateFallbackIcon(int size = 32)
    {
        using var bitmap = new Bitmap(size, size, PixelFormat.Format32bppArgb);
        using var g = Graphics.FromImage(bitmap);

        // Background gradient (blue)
        using var brush = new System.Drawing.Drawing2D.LinearGradientBrush(
            new Rectangle(0, 0, size, size),
            Color.FromArgb(14, 165, 233),  // #0ea5e9
            Color.FromArgb(3, 105, 161),   // #0369a1
            System.Drawing.Drawing2D.LinearGradientMode.ForwardDiagonal);

        // Rounded rectangle background
        var cornerRadius = size / 4;
        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
        g.FillRectangle(brush, 0, 0, size, size);

        // White chart line (simplified version of SaleCast logo)
        using var pen = new Pen(Color.White, size / 10f)
        {
            StartCap = System.Drawing.Drawing2D.LineCap.Round,
            EndCap = System.Drawing.Drawing2D.LineCap.Round,
            LineJoin = System.Drawing.Drawing2D.LineJoin.Round
        };

        var margin = size / 5;
        g.DrawLines(pen, new[]
        {
            new Point(margin, size - margin),
            new Point(size / 2, size / 2),
            new Point(size - margin, margin)
        });

        return Icon.FromHandle(bitmap.GetHicon());
    }

    /// <summary>
    /// Render SVG to PNG bytes (cross-platform, for macOS menu bar)
    /// </summary>
    public static byte[]? RenderSvgToPng(string svgPath, int size = 22)
    {
        try
        {
            using var svg = new SKSvg();
            svg.Load(svgPath);

            if (svg.Picture == null) return null;

            var bounds = svg.Picture.CullRect;
            var scale = Math.Min(size / bounds.Width, size / bounds.Height);

            using var bitmap = new SKBitmap(size, size);
            using var canvas = new SKCanvas(bitmap);

            canvas.Clear(SKColors.Transparent);
            canvas.Scale(scale);
            canvas.DrawPicture(svg.Picture);

            using var image = SKImage.FromBitmap(bitmap);
            using var data = image.Encode(SKEncodedImageFormat.Png, 100);

            return data.ToArray();
        }
        catch
        {
            return null;
        }
    }
}
