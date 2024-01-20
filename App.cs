using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

[assembly: AssemblyTitle("Simple PiP")]
[assembly: AssemblyDescription("Always-on-top Picture-in-Picture for any window")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("")]
[assembly: AssemblyProduct("Simple PiP")]
[assembly: AssemblyCopyright("Copyright © 2024")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

[assembly: ComVisible(false)]

[assembly: Guid("93afd128-87ca-43a1-8512-a56e7c8fbb52")]

[assembly: AssemblyVersion("1.0.0")]
[assembly: AssemblyFileVersion("1.0.0")]

namespace simple_pip
{
    class App
    {
        [STAThread]
        static void Main()
        {
            WinApi.SetProcessDPIAware();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }

    public class MainForm : Form
    {
        private System.ComponentModel.IContainer components = null;

        private NotifyIcon trayIcon;
        private ContextMenuStrip mainContextMenu;
        private Label instructionsLabel;

        private bool dragging;
        private int draggingStartTop;
        private int draggingStartLeft;
        private Point draggingStartScreenPoint;
        private bool dragged;

        private IntPtr selectedWindowHandle = IntPtr.Zero;
        private IntPtr thumbnailHandle = IntPtr.Zero;

        WinEventDelegate foregroundDelegate = null;
        IntPtr foregroundHook = IntPtr.Zero;

        public MainForm()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            
            this.trayIcon = new NotifyIcon(this.components);
            this.mainContextMenu = new ContextMenuStrip(this.components);
            this.instructionsLabel = new Label();

            this.SuspendLayout();

            // 
            // trayIcon
            // 
            this.trayIcon.ContextMenuStrip = this.mainContextMenu;
            this.trayIcon.Icon = Resources.MainIcon;
            this.trayIcon.Visible = true;

            // 
            // mainContextMenu
            // 
            this.mainContextMenu.AutoClose = true;
            this.mainContextMenu.Opening += new System.ComponentModel.CancelEventHandler(MainContextMenu_Opening);

            // 
            // instructionsLabel
            // 
            this.instructionsLabel.AutoSize = true;
            this.instructionsLabel.Enabled = false;
            this.instructionsLabel.ForeColor = SystemColors.GrayText;
            this.instructionsLabel.Text = "Right-click here or in tray\r\nand select a window";
            this.instructionsLabel.TextAlign = ContentAlignment.MiddleCenter;

            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new SizeF(96F, 96F);
            this.AutoScaleMode = AutoScaleMode.Dpi;
            this.ClientSize = new Size(400, 250);
            this.Controls.Add(this.instructionsLabel);
            this.FormBorderStyle = FormBorderStyle.SizableToolWindow;
            this.ControlBox = false;
            this.Icon = Resources.MainIcon;
            this.Name = "MainForm";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.TopMost = true;

            this.ResumeLayout(false);

            UpdateInstructionsLabel();

            foregroundDelegate = new WinEventDelegate(ForegroundWinEventProc);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void UpdateInstructionsLabel()
        {
            this.instructionsLabel.Top = (ClientSize.Height - this.instructionsLabel.Height) / 2;
            this.instructionsLabel.Left = (ClientSize.Width - this.instructionsLabel.Width) / 2;
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);

            UpdateInstructionsLabel();

            if (thumbnailHandle != IntPtr.Zero)
            {
                UpdateThumbnail();
            }
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);

            if (e.Button == MouseButtons.Left)
            {
                dragged = false;
                dragging = true;
                draggingStartTop = Top;
                draggingStartLeft = Left;
                draggingStartScreenPoint = PointToScreen(e.Location);
            }
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            if (dragging)
            {
                Point p = PointToScreen(e.Location);

                Top = draggingStartTop + p.Y - draggingStartScreenPoint.Y;
                Left = draggingStartLeft + p.X - draggingStartScreenPoint.X;

                if (!dragged)
                {
                    dragged = Math.Abs(p.X - draggingStartScreenPoint.X) > 2
                        || Math.Abs(p.Y - draggingStartScreenPoint.Y) > 2;
                }
            }
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);

            if (e.Button == MouseButtons.Left)
            {
                dragging = false;

                if (!dragged && selectedWindowHandle != IntPtr.Zero)
                {
                    if (WinApi.IsWindow(selectedWindowHandle))
                    {
                        //TopMost = false;
                        if (WinApi.IsIconic(selectedWindowHandle))
                        {
                            WinApi.ShowWindow(selectedWindowHandle, WinApi.SW_RESTORE);
                        }
                        WinApi.SetForegroundWindow(selectedWindowHandle);
                    }
                    else
                    {
                        Reset();
                    }
                }
            }
            else if (e.Button == MouseButtons.Right)
            {
                this.mainContextMenu.Show(this, e.Location);
            }
        }

        void MainContextMenu_Opening(object sender, System.ComponentModel.CancelEventArgs e)
        {
            WindowList.Refresh();

            mainContextMenu.Items.Clear();

            ToolStripMenuItem item;

            // https://github.com/christianrondeau/GoToWindow/blob/c0605a7c4ff84d47205da149dcc815854fc4b8c8/GoToWindow.Api/WindowsListFactory.cs
            foreach (var wnd in WindowList.Items)
            {
                if (wnd.Handle == Handle)
                {
                    continue;
                }

                if (wnd.ClassName == "Progman" || wnd.ClassName == "Windows.UI.Core.CoreWindow")
                {
                    continue;
                }

                item = new ToolStripMenuItem();

                if (wnd.Title.Length > 55)
                {
                    item.Text = wnd.Title.Substring(0, 30) + " ... " + wnd.Title.Substring(wnd.Title.Length - 20);
                    item.ToolTipText = wnd.Title;
                }
                else
                {
                    item.Text = wnd.Title;
                }

                if (
                    WinApi.SendMessageTimeout(
                        wnd.Handle,
                        0x007F,         // WM_GETICON
                        new IntPtr(2),  // ICON_SMALL2. Retrieves the small icon provided by the application.
                                        // If the application does not provide one,
                                        // the system uses the system-generated icon for that window.
                        new IntPtr(0),  // The DPI of the icon being retrieved.
                        WinApi.SendMessageTimeoutFlags.AbortIfHung | WinApi.SendMessageTimeoutFlags.Block,
                        300,
                        out IntPtr hIcon
                    ) == IntPtr.Zero)
                {
                    hIcon = IntPtr.Zero;
                }

                if (hIcon != IntPtr.Zero)
                {
                    try
                    {
                        item.Image = Icon.FromHandle(hIcon).ToBitmap();
                    }
                    catch (Exception)
                    {
                        item.Image = null;
                    }
                }

                item.Tag = wnd;
                item.Click += MenuWindow_Click;

                mainContextMenu.Items.Add(item);
            }

            mainContextMenu.Items.Add("-");

            item = new ToolStripMenuItem("Resize");
            item.Click += MenuResize_Click;
            mainContextMenu.Items.Add(item);

            item = new ToolStripMenuItem("Exit");
            item.Click += MenuExit_Click;
            mainContextMenu.Items.Add(item);

            e.Cancel = false;
        }

        private void MenuWindow_Click(object sender, EventArgs e)
        {
            var item = (ToolStripMenuItem)sender;
            var wnd = (WindowEntry)item.Tag;

            if (thumbnailHandle != IntPtr.Zero)
            {
                Thumbnail.Destroy(thumbnailHandle);
            }

            Thumbnail.Init(Handle, wnd.Handle, out thumbnailHandle);

            if (thumbnailHandle ==  IntPtr.Zero)
            {
                return;
            }

            UpdateThumbnail();
            FitThumbnailSize();

            selectedWindowHandle = wnd.Handle;

            instructionsLabel.Visible = false;

            if (foregroundHook != IntPtr.Zero)
            {
                WinApi.UnhookWinEvent(foregroundHook);
            }

            foregroundHook = WinApi.SetWinEventHook(
                WinApi.EVENT_SYSTEM_FOREGROUND,
                WinApi.EVENT_SYSTEM_FOREGROUND,
                IntPtr.Zero,
                foregroundDelegate,
                0,
                0,
                WinApi.WINEVENT_OUTOFCONTEXT | WinApi.WINEVENT_SKIPOWNPROCESS
            );
        }

        private void MenuResize_Click(object sender, EventArgs e)
        {
            FitThumbnailSize();
        }

        private void MenuExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        public void ForegroundWinEventProc(IntPtr hWinEventHook, uint eventType, IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
        {
            if (eventType == WinApi.EVENT_SYSTEM_FOREGROUND)
            {
                if (selectedWindowHandle != IntPtr.Zero)
                {
                    // Using hwnd is not reliable here because of how Alt-Tab works.
                    // Sometimes the latest hwnd belongs to explorer.exe instead of the actual foreground window.
                    TopMost = WinApi.GetForegroundWindow() != selectedWindowHandle;

                    if (WinApi.IsWindow(selectedWindowHandle))
                    {
                        if (!TopMost)
                        {
                            WinApi.SetWindowPos(selectedWindowHandle, 0, 0, 0, 0, 0, 3);
                        }
                    }
                    else {
                        Reset();
                    }
                }
                else
                {
                    Reset();
                }
            }
        }

        private void Reset()
        {
            if (thumbnailHandle != IntPtr.Zero)
            {
                Thumbnail.Destroy(thumbnailHandle);
                thumbnailHandle = IntPtr.Zero;
            }

            if (foregroundHook != IntPtr.Zero)
            {
                WinApi.UnhookWinEvent(foregroundHook);
                foregroundHook = IntPtr.Zero;
            }

            selectedWindowHandle = IntPtr.Zero;

            TopMost = true;

            instructionsLabel.Visible = true;
            UpdateInstructionsLabel();
        }

        private void FitThumbnailSize()
        {
            if (thumbnailHandle == IntPtr.Zero)
            {
                return;
            }

            ClientSize = CalculateThumbnailSize();
        }

        private void UpdateThumbnail()
        {
            if (thumbnailHandle == IntPtr.Zero)
            {
                return;
            }

            var thumbSize = CalculateThumbnailSize();

            var thumbX = (ClientRectangle.Width - thumbSize.Width) / 2;
            var thumbY = (ClientRectangle.Height - thumbSize.Height) / 2;

            Thumbnail.Update(thumbnailHandle, new Rectangle(thumbX, thumbY, thumbSize.Width, thumbSize.Height));
        }

        private Size CalculateThumbnailSize()
        {
            var sourceSize = Thumbnail.GetSourceSize(thumbnailHandle);
            var sourceRatio = (double)sourceSize.Width / sourceSize.Height;

            var canvasRatio = (double)ClientRectangle.Width / ClientRectangle.Height;

            return canvasRatio > sourceRatio
                ? new Size((int)(ClientRectangle.Height * sourceRatio), ClientRectangle.Height)
                : new Size(ClientRectangle.Width, (int)(ClientRectangle.Width / sourceRatio));
        }
    }

    class WindowEntry
    {
        public IntPtr Handle { get; set; }
        public string Title { get; set; }
        public string ClassName { get; set; }

        public WindowEntry(IntPtr hwnd, string title, string className)
        {
            Handle = hwnd;
            Title = title;
            ClassName = className;
        }
    }

    static class WindowList
    {
        private static List<WindowEntry> list = new List<WindowEntry>();

        public static List<WindowEntry> Items { get { return list; } }

        public static void Refresh()
        {
            list.Clear();
            WinApi.EnumWindows(AddWindowListEntry, IntPtr.Zero);
        }

        private static bool AddWindowListEntry(IntPtr hwnd, IntPtr lParam)
        {
            if (!WinApi.IsWindowVisible(hwnd))
            {
                return true;
            }

            StringBuilder sb;

            string title;

            int length = WinApi.GetWindowTextLength(hwnd);

            if (length > 0)
            {
                sb = new StringBuilder(length + 1);
                title = WinApi.GetWindowText(hwnd, sb, sb.Capacity) > 0 ? sb.ToString() : string.Empty;
            }
            else
            {
                title = string.Empty;
            }

            if (string.IsNullOrEmpty(title))
            {
                return true;
            }

            sb = new StringBuilder(100);
            var className = WinApi.GetClassName(hwnd, sb, sb.Capacity) > 0 ? sb.ToString() : string.Empty;

            list.Add(new WindowEntry(hwnd, title, className));

            return true;
        }
    }

    static class Thumbnail
    {
        public static void Init(IntPtr dest, IntPtr src, out IntPtr thumb)
        {
            var result = WinApi.DwmRegisterThumbnail(dest, src, out thumb);

            if (result != 0 || thumb == IntPtr.Zero)
            {
                throw new Exception("Cannot initialize thumbnail");
            }
        }

        public static void Update(IntPtr thumb, Rectangle destRect)
        {
            if (thumb == IntPtr.Zero)
            {
                return;
            }

            var dest = new WinApi.Rect(
                destRect.X, 
                destRect.Y,
                destRect.X + destRect.Width,
                destRect.Y + destRect.Height
            );

            var props = new WinApi.DwmThumbnailProperties
            {
                dwFlags = (int)(
                    WinApi.DWM_TNP.DWM_TNP_VISIBLE 
                    | WinApi.DWM_TNP.DWM_TNP_RECTDESTINATION 
                    | WinApi.DWM_TNP.DWM_TNP_OPACITY
                ),
                fVisible = true,
                opacity = 255,
                rcDestination = dest
            };

            WinApi.DwmUpdateThumbnailProperties(thumb, ref props);
        }

        public static Size GetSourceSize(IntPtr thumb)
        {
            if (thumb == IntPtr.Zero)
            {
                throw new Exception("Cannot get source size for unregistered thumbnail");
            }
            WinApi.DwmQueryThumbnailSourceSize(thumb, out var pSourceSize);
            return new Size(pSourceSize.x, pSourceSize.y);
        }

        public static void Destroy(IntPtr thumb)
        {
            if (thumb == IntPtr.Zero)
            {
                return;
            }
            WinApi.DwmUnregisterThumbnail(thumb);
        }
    }

    delegate void WinEventDelegate(
        IntPtr hWinEventHook,
        uint eventType,
        IntPtr hwnd,
        int idObject,
        int idChild,
        uint dwEventThread,
        uint dwmsEventTime
    );

    class WinApi
    {

        [Flags]
        public enum SendMessageTimeoutFlags : uint
        {
            AbortIfHung = 2,
            Block = 1,
            Normal = 0
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = false)]
        public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern IntPtr SendMessageTimeout(IntPtr hwnd, uint message, IntPtr wparam, IntPtr lparam, SendMessageTimeoutFlags flags, uint timeout, out IntPtr result);

        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport("user32.dll", SetLastError = false)]
        public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        public const int WINEVENT_OUTOFCONTEXT = 0;
        public const int WINEVENT_SKIPOWNTHREAD = 1;
        public const int WINEVENT_SKIPOWNPROCESS = 2;
        public const int WINEVENT_INCONTEXT = 4;

        public const uint EVENT_SYSTEM_FOREGROUND = 3;

        public const int OBJID_WINDOW = 0;
        public const int CHILDID_SELF = 0;

        [DllImport("user32.dll")]
        public static extern IntPtr SetWinEventHook(uint eventMin, uint eventMax, IntPtr hmodWinEventProc, WinEventDelegate lpfnWinEventProc, uint idProcess, uint idThread, uint dwFlags);

        [DllImport("user32.dll")]
        public static extern bool UnhookWinEvent(IntPtr hWinEventHook);

        [return: MarshalAs(UnmanagedType.Bool)]
        public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindowVisible(IntPtr hWnd);

        public const uint SW_SHOW = 0x05;
        public const uint SW_RESTORE = 0x09;

        [DllImport("user32.dll")]
        public static extern int ShowWindow(IntPtr hWnd, uint Msg);

        [DllImport("user32.dll")]
        public static extern bool IsIconic(IntPtr handle);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern int GetWindowText(IntPtr hWnd, [Out] StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll", EntryPoint = "SetWindowPos")]
        public static extern IntPtr SetWindowPos(IntPtr hWnd, int hWndInsertAfter, int x, int Y, int cx, int cy, int wFlags);

        [StructLayout(LayoutKind.Sequential)]
        public struct Psize
        {
            public int x;
            public int y;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct Rect
        {
            public int Left, Top, Right, Bottom;

            public Rect(int left, int top, int right, int bottom)
            {
                Left = left;
                Top = top;
                Right = right;
                Bottom = bottom;
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct DwmThumbnailProperties
        {
            public int dwFlags;
            public Rect rcDestination;
            public Rect rcSource;
            public byte opacity;
            public bool fVisible;
            public bool fSourceClientAreaOnly;
        }

        public enum DWM_TNP : int
        {
            DWM_TNP_RECTDESTINATION = 0x1,
            DWM_TNP_RECTSOURCE = 0x2,
            DWM_TNP_VISIBLE = 0x8,
            DWM_TNP_OPACITY = 0x4,
            DWM_TNP_SOURCECLIENTAREAONLY = 0x10
        }

        [DllImport("dwmapi.dll")]
        public static extern int DwmRegisterThumbnail(IntPtr dest, IntPtr src, out IntPtr thumb);

        [DllImport("dwmapi.dll")]
        public static extern int DwmUnregisterThumbnail(IntPtr thumb);

        [DllImport("dwmapi.dll")]
        public static extern int DwmQueryThumbnailSourceSize(IntPtr thumb, out Psize size);

        [DllImport("dwmapi.dll")]
        public static extern int DwmUpdateThumbnailProperties(IntPtr hThumb, ref DwmThumbnailProperties props);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern bool SetProcessDPIAware();
    }

    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "17.0.0.0")]
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    internal class Resources {
        private static global::System.Resources.ResourceManager resourceMan;
        
        private static global::System.Globalization.CultureInfo resourceCulture;
        
        [global::System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        internal Resources() {
        }
        
        /// <summary>
        ///   Returns the cached ResourceManager instance used by this class.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Resources.ResourceManager ResourceManager {
            get {
                if (object.ReferenceEquals(resourceMan, null)) {
                    global::System.Resources.ResourceManager temp = new global::System.Resources.ResourceManager("simple_pip.Resources", typeof(Resources).Assembly);
                    resourceMan = temp;
                }
                return resourceMan;
            }
        }
        
        /// <summary>
        ///   Overrides the current thread's CurrentUICulture property for all
        ///   resource lookups using this strongly typed resource class.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Globalization.CultureInfo Culture {
            get {
                return resourceCulture;
            }
            set {
                resourceCulture = value;
            }
        }
        
        /// <summary>
        ///   Looks up a localized resource of type System.Drawing.Icon similar to (Icon).
        /// </summary>
        internal static System.Drawing.Icon MainIcon {
            get {
                object obj = ResourceManager.GetObject("MainIcon", resourceCulture);
                return ((System.Drawing.Icon)(obj));
            }
        }
    }
}
