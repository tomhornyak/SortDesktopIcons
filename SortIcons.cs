using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace SortDesktopIcons
{

    static class SortIcons
    {
        [ComImport, Guid("6D5140C1-7436-11CE-8034-00AA006009FA"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IServiceProvider
        {
            [return: MarshalAs(UnmanagedType.IUnknown)]
            object QueryService([MarshalAs(UnmanagedType.LPStruct)] Guid service, [MarshalAs(UnmanagedType.LPStruct)] Guid riid);
        }

        // note: for the following interfaces, not all methods are defined as we don't use them here
        [ComImport, Guid("000214E2-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IShellBrowser
        {
            void _VtblGap1_12(); // skip 12 methods https://stackoverflow.com/a/47567206/403671

            [return: MarshalAs(UnmanagedType.IUnknown)]
            object QueryActiveShellView();
        }

        [ComImport, Guid("cde725b0-ccc9-4519-917e-325d72fab4ce"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IFolderView
        {
            void _VtblGap1_3(); // skip 3 methods

            IntPtr Item(int iItemIndex);
            int ItemCount(uint uFlags = 0);

            void _VtblGap2_3(); // skip 2 methods

            void GetItemPosition(IntPtr pidl, out POINT ppt);

            void _VtblGap1_4(); // skip 4 methods

            void SelectAndPositionItems(int cidl, [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] IntPtr[] apidl, [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] POINT[] apt, SVSIF dwFlags);

            // more undefined methods
        }

        public enum FOLDERFLAGS
        {
            FWF_AUTOARRANGE = 0x1,
            FWF_SNAPTOGRID = 0x4
        }
        [ComImport, Guid("1af3a467-214f-4298-908e-06b03e0b39f9"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IFolderView2
        {
            //void _VtblGap1_26(); // skip 14 (IFolderView) + 12 methods
            void _VtblGap1_21(); // skip 14 (IFolderView) + 12 methods
            uint SetCurrentFolderFlags(int dwMask, int dwFlags);
            void _VtblGap1_4();
            IShellItem GetItem(int iItemIndex, [MarshalAs(UnmanagedType.LPStruct)] Guid riid);

            // more undefined methods
        }

        [ComImport, Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IShellItem
        {
            [return: MarshalAs(UnmanagedType.IUnknown)]
            object BindToHandler(System.Runtime.InteropServices.ComTypes.IBindCtx pbc, [MarshalAs(UnmanagedType.LPStruct)] Guid bhid, [MarshalAs(UnmanagedType.LPStruct)] Guid riid);

            IShellItem GetParent();

            [return: MarshalAs(UnmanagedType.LPWStr)]
            string GetDisplayName(SIGDN sigdnName);

            [return: MarshalAs(UnmanagedType.U4)]
            uint GetAttributesOf(uint sfgaoMask);

            // more undefined methods
        }

        public struct POINT
        {
            public int x;
            public int y;

            //Eliminate stupit warning about uninitialized y
            public POINT()
            {
                x = -1;
                y = -1;
            }
            public override string ToString()
            {
                return $"({x},{y})";
            }
        }

        public enum SIGDN
        {
            SIGDN_NORMALDISPLAY,
            SIGDN_PARENTRELATIVEPARSING,
            SIGDN_DESKTOPABSOLUTEPARSING,
            SIGDN_PARENTRELATIVEEDITING,
            SIGDN_DESKTOPABSOLUTEEDITING,
            SIGDN_FILESYSPATH,
            SIGDN_URL,
            SIGDN_PARENTRELATIVEFORADDRESSBAR,
            SIGDN_PARENTRELATIVE,
            SIGDN_PARENTRELATIVEFORUI
        }

        [Flags]
        public enum SVSIF
        {
            SVSI_DESELECT = 0,
            SVSI_SELECT = 0x1,
            SVSI_EDIT = 0x3,
            SVSI_DESELECTOTHERS = 0x4,
            SVSI_ENSUREVISIBLE = 0x8,
            SVSI_FOCUSED = 0x10,
            SVSI_TRANSLATEPT = 0x20,
            SVSI_SELECTIONMARK = 0x40,
            SVSI_POSITIONITEM = 0x80,
            SVSI_CHECK = 0x100,
            SVSI_CHECK2 = 0x200,
            SVSI_KEYBOARDSELECT = 0x401,
            SVSI_NOTAKEFOCUS = 0x40000000
        }
        public enum SFGAO
        {
            CANCOPY = 0x1,
            CANMOVE = 0x2,
            CANLINK = 0x4,
            STORAGE = 0x8,
            CANRENAME = 0x10,
            CANDELETE = 0x20,
            HASPROPSHEET = 0x40,
            DROPTARGET = 0x100,
            CAPABILITYMASK = 0x177,
            ENCRYPTED = 0x2000,
            ISSLOW = 0x4000,
            GHOSTED = 0x8000,
            LINK = 0x10000,
            SHARE = 0x20000,
            READONLY = 0x40000,
            HIDDEN = 0x80000,
            DISPLAYATTRMASK = 0xFC000,
            STREAM = 0x400000,
            STORAGEANCESTOR = 0x800000,
            VALIDATE = 0x1000000,
            REMOVABLE = 0x2000000,
            COMPRESSED = 0x4000000,
            BROWSABLE = 0x8000000,
            FILESYSANCESTOR = 0x10000000,
            FOLDER = 0x20000000,
            FILESYSTEM = 0x40000000,
            HASSUBFOLDER = unchecked((int)0x80000000),
            CONTENTSMASK = unchecked((int)0x80000000),
            STORAGECAPMASK = 0x70C50008,
            PKEYSFGAOMASK = unchecked((int)0x81044000)
        }
        struct icon
        {
            public int index;
            public string name;
            public POINT locn;
            public IntPtr pid;
            public string type;

            public override string ToString()
            {
                return $"Index: {index}, {type}, Name: {name}, Location: {locn}";
            }
        }

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
        [DllImport("Kernel32.dll", CallingConvention = CallingConvention.StdCall, SetLastError = true)]
        private static extern IntPtr GetConsoleWindow();

        [DllImport("User32.dll", CallingConvention = CallingConvention.StdCall, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool ShowWindow([In] IntPtr hWnd, [In] Int32 nCmdShow);

        const string SHELL_CLASS_NAME = "Progman";
        const uint WM_COMMAND = 0x111;
        const int CMD_SORT_BY_NAME = 0x800020;

        [STAThread]
        public static void Main()
        {
            IntPtr handle = GetConsoleWindow();
            ShowWindow(handle, 6);

            IntPtr desktopHandle = FindWindow(SHELL_CLASS_NAME, null);

            dynamic app = Activator.CreateInstance(Type.GetTypeFromProgID("Shell.Application"));
            var windows = app.Windows;

            const int SWC_DESKTOP = 8;
            const int SWFO_NEEDDISPATCH = 1;
            var hwnd = 0;
            var disp = windows.FindWindowSW(Type.Missing, Type.Missing, SWC_DESKTOP, ref hwnd, SWFO_NEEDDISPATCH);

            var sp = (IServiceProvider)disp;
            var SID_STopLevelBrowser = new Guid("4c96be40-915c-11cf-99d3-00aa004ae837");

            var browser = (IShellBrowser)sp.QueryService(SID_STopLevelBrowser, typeof(IShellBrowser).GUID);
            var view = (IFolderView)browser.QueryActiveShellView();
            var view2 = (IFolderView2)view;

            view2.SetCurrentFolderFlags((int)(FOLDERFLAGS.FWF_SNAPTOGRID | FOLDERFLAGS.FWF_AUTOARRANGE),
                (int)(FOLDERFLAGS.FWF_SNAPTOGRID | FOLDERFLAGS.FWF_AUTOARRANGE));
            if (desktopHandle != IntPtr.Zero)
            {
                SendMessage(desktopHandle, WM_COMMAND, new IntPtr(CMD_SORT_BY_NAME), IntPtr.Zero);
                Console.WriteLine("Desktop icons sorted by name.");
            }
            view2.SetCurrentFolderFlags((int)(FOLDERFLAGS.FWF_AUTOARRANGE), 0);
            var icons = new List<icon>();

            // get all items, dump & sets their position (here y+= 150)
            for (var i = 0; i < view.ItemCount(); i++)
            {
                // get some item's info to be able to determine if we want to move it or not
                var item = view2.GetItem(i, typeof(IShellItem).GUID);
                uint attrib = (uint)SFGAO.FOLDER;
                var type = item.GetAttributesOf(attrib) == 0 ? "File" : "Folder";
                var displayName = item.GetDisplayName(SIGDN.SIGDN_DESKTOPABSOLUTEPARSING);

                var pidl = view.Item(i);
                view.GetItemPosition(pidl, out var pt);

                var ic = new icon { index = i, locn = pt, name = displayName, pid = pidl, type = type };
                icons.Add(ic);
                Console.WriteLine(ic.ToString());
            }
            var byPos = new List<icon>();
            byPos =
                icons.OrderBy(icon => icon.locn.x)
                .ThenBy(icon => icon.locn.y)
                .ToList();

            var sortedIcons = new List<icon>();
            sortedIcons =
                byPos.OrderByDescending(icon => icon.type)
                .ThenBy(icon => icon.name)
                .ToList();
            var ptrs = new List<IntPtr>();
            var locns = new List<POINT>();
            for (int i = 0; i < sortedIcons.Count(); i++)
            {
                var tempIcon = sortedIcons[i]; // Create a temporary variable
                tempIcon.locn = byPos[i].locn; // Modify the temporary variable
                sortedIcons[i] = tempIcon;
                ptrs.Add(tempIcon.pid);
                locns.Add(tempIcon.locn);
            }
            IntPtr[] apidl = ptrs.ToArray();
            POINT[] apt = locns.ToArray();

            view.SelectAndPositionItems(apidl.Length, apidl, apt, SVSIF.SVSI_POSITIONITEM);

        }

    }
}

