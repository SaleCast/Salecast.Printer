using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using SkiaSharp;
using Svg.Skia;

namespace SaleCast.Printer.Services;

/// <summary>
/// macOS menu bar (NSStatusItem) service using native Cocoa API via ObjC runtime interop.
/// Provides a clickable status bar icon with context menu, equivalent to Windows TrayIconService.
/// </summary>
[SupportedOSPlatform("macos")]
public class MacMenuBarService : IDisposable
{
    private readonly ILogger<MacMenuBarService> _logger;
    private readonly IHostApplicationLifetime _appLifetime;
    private readonly int _port;
    private bool _disposed;
    private bool _initialized;
    private volatile bool _running;
    private IntPtr _statusItem;
    private IntPtr _delegateInstance;

    // Must keep strong references to prevent GC of native callbacks
    private static readonly List<Delegate> s_prevent_gc = new();

    #region ObjC Runtime P/Invoke

    private const string ObjCRuntime = "/usr/lib/libobjc.dylib";
    private const string AppKitPath = "/System/Library/Frameworks/AppKit.framework/AppKit";

    [DllImport("/usr/lib/libdl.dylib")]
    private static extern IntPtr dlopen(string path, int mode);

    [DllImport(ObjCRuntime)]
    private static extern IntPtr objc_getClass(string className);

    [DllImport(ObjCRuntime)]
    private static extern IntPtr sel_registerName(string selectorName);

    [DllImport(ObjCRuntime, EntryPoint = "objc_msgSend")]
    private static extern IntPtr objc_msgSend(IntPtr receiver, IntPtr selector);

    [DllImport(ObjCRuntime, EntryPoint = "objc_msgSend")]
    private static extern IntPtr objc_msgSend(IntPtr receiver, IntPtr selector, IntPtr arg1);

    [DllImport(ObjCRuntime, EntryPoint = "objc_msgSend")]
    private static extern IntPtr objc_msgSend(IntPtr receiver, IntPtr selector, IntPtr arg1, IntPtr arg2);

    [DllImport(ObjCRuntime, EntryPoint = "objc_msgSend")]
    private static extern IntPtr objc_msgSend(IntPtr receiver, IntPtr selector, IntPtr arg1, IntPtr arg2, IntPtr arg3);

    [DllImport(ObjCRuntime, EntryPoint = "objc_msgSend")]
    private static extern IntPtr objc_msgSend_double(IntPtr receiver, IntPtr selector, double arg1);

    [DllImport(ObjCRuntime, EntryPoint = "objc_msgSend")]
    private static extern void objc_msgSend_void(IntPtr receiver, IntPtr selector);

    [DllImport(ObjCRuntime, EntryPoint = "objc_msgSend")]
    private static extern void objc_msgSend_void(IntPtr receiver, IntPtr selector, IntPtr arg1);

    [DllImport(ObjCRuntime, EntryPoint = "objc_msgSend")]
    private static extern void objc_msgSend_void_bool(IntPtr receiver, IntPtr selector, [MarshalAs(UnmanagedType.I1)] bool arg1);

    [DllImport(ObjCRuntime, EntryPoint = "objc_msgSend")]
    private static extern void objc_msgSend_void_size(IntPtr receiver, IntPtr selector, NSSize arg1);

    [DllImport(ObjCRuntime)]
    private static extern IntPtr objc_allocateClassPair(IntPtr superclass, string name, int extraBytes);

    [DllImport(ObjCRuntime)]
    private static extern void objc_registerClassPair(IntPtr cls);

    [DllImport(ObjCRuntime)]
    private static extern bool class_addMethod(IntPtr cls, IntPtr sel, IntPtr imp, string types);

    [StructLayout(LayoutKind.Sequential)]
    private struct NSSize
    {
        public double Width;
        public double Height;
    }

    private delegate void ObjCMethodDelegate(IntPtr self, IntPtr sel, IntPtr sender);

    #endregion

    public MacMenuBarService(
        ILogger<MacMenuBarService> logger,
        IHostApplicationLifetime appLifetime,
        IConfiguration configuration)
    {
        _logger = logger;
        _appLifetime = appLifetime;
        _port = configuration.GetValue("Api:Port", 5123);
    }

    /// <summary>
    /// Initialize the NSApplication, status bar icon, and context menu.
    /// Must be called on the main thread before Run().
    /// </summary>
    public void Initialize()
    {
        try
        {
            // Ensure AppKit framework is loaded
            dlopen(AppKitPath, 1); // RTLD_LAZY

            // Initialize NSApplication
            var nsApp = objc_msgSend(objc_getClass("NSApplication"), sel_registerName("sharedApplication"));

            // Hide from Dock (NSApplicationActivationPolicyAccessory = 1)
            objc_msgSend(nsApp, sel_registerName("setActivationPolicy:"), (IntPtr)1);

            // Create status bar item with variable length
            var statusBar = objc_msgSend(objc_getClass("NSStatusBar"), sel_registerName("systemStatusBar"));
            _statusItem = objc_msgSend_double(statusBar, sel_registerName("statusItemWithLength:"), -1.0);
            objc_msgSend(_statusItem, sel_registerName("retain"));

            // Configure the button (icon or text fallback)
            var button = objc_msgSend(_statusItem, sel_registerName("button"));
            if (button != IntPtr.Zero)
            {
                if (!TrySetIcon(button))
                {
                    objc_msgSend_void(button, sel_registerName("setTitle:"), CreateNSString("SC"));
                }
            }

            // Register ObjC delegate class for menu actions
            RegisterDelegateClass();

            // Create and attach context menu
            CreateMenu();

            _initialized = true;
            _logger.LogInformation("macOS menu bar icon initialized");
            _logger.LogInformation("API available at: http://localhost:{Port}", _port);
            _logger.LogInformation("Swagger UI: http://localhost:{Port}/swagger", _port);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to initialize macOS menu bar icon");
            throw;
        }
    }

    /// <summary>
    /// Run the macOS event loop. Blocks until Quit is called.
    /// Must be called on the main thread.
    /// </summary>
    public void Run()
    {
        if (!_initialized)
        {
            _logger.LogWarning("Menu bar not initialized, falling back to blocking wait");
            _appLifetime.ApplicationStopping.WaitHandle.WaitOne();
            return;
        }

        _running = true;
        var nsApp = objc_msgSend(objc_getClass("NSApplication"), sel_registerName("sharedApplication"));
        objc_msgSend_void(nsApp, sel_registerName("finishLaunching"));

        var runLoop = objc_msgSend(objc_getClass("NSRunLoop"), sel_registerName("currentRunLoop"));
        var mode = CreateNSString("kCFRunLoopDefaultMode");

        while (_running)
        {
            // Autorelease pool to drain autoreleased ObjC objects each iteration
            var pool = objc_msgSend(
                objc_msgSend(objc_getClass("NSAutoreleasePool"), sel_registerName("alloc")),
                sel_registerName("init"));

            var date = objc_msgSend_double(
                objc_getClass("NSDate"),
                sel_registerName("dateWithTimeIntervalSinceNow:"), 1.0);
            objc_msgSend(runLoop, sel_registerName("runMode:beforeDate:"), mode, date);

            objc_msgSend_void(pool, sel_registerName("drain"));
        }
    }

    #region Icon

    private bool TrySetIcon(IntPtr button)
    {
        try
        {
            var svgPath = Path.Combine(AppContext.BaseDirectory, "Assets", "icon.svg");
            if (!File.Exists(svgPath)) return false;

            var pngData = RenderSvgToPng(svgPath, 44); // 22pt * 2 for Retina
            if (pngData == null) return false;

            var pinned = GCHandle.Alloc(pngData, GCHandleType.Pinned);
            try
            {
                var nsData = objc_msgSend(
                    objc_msgSend(objc_getClass("NSData"), sel_registerName("alloc")),
                    sel_registerName("initWithBytes:length:"),
                    pinned.AddrOfPinnedObject(), (IntPtr)pngData.Length);

                var nsImage = objc_msgSend(
                    objc_msgSend(objc_getClass("NSImage"), sel_registerName("alloc")),
                    sel_registerName("initWithData:"), nsData);

                // NSData no longer needed, NSImage copies it
                objc_msgSend_void(nsData, sel_registerName("release"));

                if (nsImage == IntPtr.Zero) return false;

                // Set logical size to 18x18 (standard menu bar icon size)
                objc_msgSend_void_size(nsImage, sel_registerName("setSize:"),
                    new NSSize { Width = 18, Height = 18 });

                // Button retains the image
                objc_msgSend_void(button, sel_registerName("setImage:"), nsImage);
                objc_msgSend_void(nsImage, sel_registerName("release"));
                return true;
            }
            finally
            {
                pinned.Free();
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to load menu bar icon");
            return false;
        }
    }

    private static byte[]? RenderSvgToPng(string svgPath, int size)
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

    #endregion

    #region Menu

    private void RegisterDelegateClass()
    {
        var superClass = objc_getClass("NSObject");
        var delegateClass = objc_allocateClassPair(superClass, "SaleCastMenuDelegate", 0);

        if (delegateClass == IntPtr.Zero)
        {
            // Class already exists (process reuse)
            delegateClass = objc_getClass("SaleCastMenuDelegate");
        }
        else
        {
            AddObjCMethod(delegateClass, "openSwagger:", OpenSwagger);
            AddObjCMethod(delegateClass, "openLogs:", OpenLogsFolder);
            AddObjCMethod(delegateClass, "openPrinters:", OpenPrintersList);
            AddObjCMethod(delegateClass, "quitApp:", Quit);
            objc_registerClassPair(delegateClass);
        }

        _delegateInstance = objc_msgSend(
            objc_msgSend(delegateClass, sel_registerName("alloc")),
            sel_registerName("init"));
        objc_msgSend(_delegateInstance, sel_registerName("retain"));
    }

    private static void AddObjCMethod(IntPtr cls, string selectorName, Action action)
    {
        ObjCMethodDelegate del = (self, sel, sender) => action();
        s_prevent_gc.Add(del);
        var imp = Marshal.GetFunctionPointerForDelegate(del);
        class_addMethod(cls, sel_registerName(selectorName), imp, "v@:@");
    }

    private void CreateMenu()
    {
        var menu = objc_msgSend(
            objc_msgSend(objc_getClass("NSMenu"), sel_registerName("alloc")),
            sel_registerName("init"));

        var version = typeof(MacMenuBarService).Assembly.GetName().Version?.ToString(3) ?? "1.0.0";
        AddMenuItem(menu, $"SaleCast Printer v{version}", null, enabled: false);
        AddSeparator(menu);
        AddMenuItem(menu, "Open API Documentation", "openSwagger:");
        AddMenuItem(menu, "Open Logs Folder", "openLogs:");
        AddSeparator(menu);
        AddMenuItem(menu, "View Printers (JSON)", "openPrinters:");
        AddSeparator(menu);
        AddMenuItem(menu, "Quit", "quitApp:", keyEquivalent: "q");

        // Status item retains the menu
        objc_msgSend_void(_statusItem, sel_registerName("setMenu:"), menu);
        objc_msgSend_void(menu, sel_registerName("release"));
    }

    private void AddMenuItem(IntPtr menu, string title, string? action,
        bool enabled = true, string keyEquivalent = "")
    {
        var nsTitle = CreateNSString(title);
        var nsKey = CreateNSString(keyEquivalent);
        var sel = action != null ? sel_registerName(action) : IntPtr.Zero;

        var item = objc_msgSend(
            objc_msgSend(objc_getClass("NSMenuItem"), sel_registerName("alloc")),
            sel_registerName("initWithTitle:action:keyEquivalent:"),
            nsTitle, sel, nsKey);

        if (action != null && _delegateInstance != IntPtr.Zero)
        {
            objc_msgSend_void(item, sel_registerName("setTarget:"), _delegateInstance);
        }

        if (!enabled)
        {
            objc_msgSend_void_bool(item, sel_registerName("setEnabled:"), false);
        }

        objc_msgSend_void(menu, sel_registerName("addItem:"), item);
    }

    private static void AddSeparator(IntPtr menu)
    {
        var separator = objc_msgSend(objc_getClass("NSMenuItem"), sel_registerName("separatorItem"));
        objc_msgSend_void(menu, sel_registerName("addItem:"), separator);
    }

    #endregion

    #region Menu Actions

    private void OpenSwagger()
    {
        var url = $"http://localhost:{_port}/swagger";
        _logger.LogInformation("Opening Swagger UI: {Url}", url);
        Process.Start(new ProcessStartInfo { FileName = "open", Arguments = url, UseShellExecute = false });
    }

    private void OpenLogsFolder()
    {
        var logsPath = Path.Combine(AppContext.BaseDirectory, "logs");
        Directory.CreateDirectory(logsPath);
        _logger.LogInformation("Opening logs folder: {Path}", logsPath);
        Process.Start(new ProcessStartInfo { FileName = "open", Arguments = logsPath, UseShellExecute = false });
    }

    private void OpenPrintersList()
    {
        var url = $"http://localhost:{_port}/printers";
        _logger.LogInformation("Opening printers list: {Url}", url);
        Process.Start(new ProcessStartInfo { FileName = "open", Arguments = url, UseShellExecute = false });
    }

    private void Quit()
    {
        _logger.LogInformation("Quit requested from menu bar");
        _running = false;
        _appLifetime.StopApplication();
    }

    #endregion

    #region Helpers

    private static IntPtr CreateNSString(string str)
    {
        var utf8 = System.Text.Encoding.UTF8.GetBytes(str);
        var pinned = GCHandle.Alloc(utf8, GCHandleType.Pinned);
        try
        {
            return objc_msgSend(
                objc_msgSend(objc_getClass("NSString"), sel_registerName("alloc")),
                sel_registerName("initWithBytes:length:encoding:"),
                pinned.AddrOfPinnedObject(), (IntPtr)utf8.Length, (IntPtr)4); // NSUTF8StringEncoding
        }
        finally
        {
            pinned.Free();
        }
    }

    #endregion

    public void Dispose()
    {
        if (_disposed) return;

        if (_statusItem != IntPtr.Zero)
        {
            var statusBar = objc_msgSend(objc_getClass("NSStatusBar"), sel_registerName("systemStatusBar"));
            objc_msgSend_void(statusBar, sel_registerName("removeStatusItem:"), _statusItem);
            objc_msgSend_void(_statusItem, sel_registerName("release"));
            _statusItem = IntPtr.Zero;
        }

        if (_delegateInstance != IntPtr.Zero)
        {
            objc_msgSend_void(_delegateInstance, sel_registerName("release"));
            _delegateInstance = IntPtr.Zero;
        }

        _disposed = true;
        GC.SuppressFinalize(this);
    }
}
