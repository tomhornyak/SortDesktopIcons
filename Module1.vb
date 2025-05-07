
Imports System.Runtime.InteropServices

Module Module1
    Public Enum HRESULT As Integer
        S_OK = 0
        S_FALSE = 1
        E_NOINTERFACE = &H80004002
        E_NOTIMPL = &H80004001
        E_FAIL = &H80004005
        E_UNEXPECTED = &H8000FFFF
        E_OUTOFMEMORY = &H8007000E
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure POINT
        Public X As Integer
        Public Y As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure RECT
        Public left As Integer
        Public top As Integer
        Public right As Integer
        Public bottom As Integer
        Public Sub New(Left As Integer, Top As Integer, Right As Integer, Bottom As Integer)
            Left = Left
            Top = Top
            Right = Right
            Bottom = Bottom
        End Sub
    End Structure

    Public Const SWC_DESKTOP = &H8

    Public Enum ShellWindowFindWindowOptions As Integer
        SWFO_NEEDDISPATCH = &H1
        SWFO_INCLUDEPENDING = &H2
        SWFO_COOKIEPASSED = &H4
    End Enum

    <DllImport("Shlwapi.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Public Function IUnknown_QueryService(ByVal punk As IntPtr, ByRef guidService As Guid, ByRef riid As Guid, <Out> ByRef ppvOut As IntPtr) As HRESULT
    End Function

    <ComImport>
    <Guid("00000114-0000-0000-C000-000000000046")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Interface IOleWindow
        Function GetWindow(<Out> ByRef phwnd As IntPtr) As HRESULT
        Function ContextSensitiveHelp(fEnterMode As Boolean) As HRESULT
    End Interface

    <ComImport>
    <Guid("000214E2-0000-0000-C000-000000000046")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Interface IShellBrowser
        Inherits IOleWindow

        Overloads Function GetWindow(<Out> ByRef phwnd As IntPtr) As HRESULT
        Overloads Function ContextSensitiveHelp(fEnterMode As Boolean) As HRESULT
        Function InsertMenusSB(IntPtrShared As IntPtr, ByRef lpMenuWidths As OLEMENUGROUPWIDTHS) As HRESULT
        Function SetMenuSB(IntPtrShared As IntPtr, holemenuRes As IntPtr, IntPtrActiveObject As IntPtr) As HRESULT
        Function RemoveMenusSB(IntPtrShared As IntPtr) As HRESULT
        Function SetStatusTextSB(pszStatusText As String) As HRESULT
        Function EnableModelessSB(fEnable As Boolean) As HRESULT
        Function TranslateAcceleratorSB(pmsg As MSG, wID As UInt16) As HRESULT
        Function BrowseObject(pidl As IntPtr, wFlags As UInteger) As HRESULT
        Function GetViewStateStream(grfMode As UInteger, <Out> ByRef ppStrm As ComTypes.IStream) As HRESULT
        Function GetControlWindow(id As UInteger, <Out> ByRef pIntPtr As IntPtr) As HRESULT
        Function SendControlMsg(id As UInteger, uMsg As UInteger, wParam As Integer, lParam As IntPtr, <Out> ByRef pret As IntPtr) As HRESULT
        Function QueryActiveShellView(<Out> ByRef ppshv As IShellView) As HRESULT
        Function OnViewWindowActive(pshv As IShellView) As HRESULT
        Function SetToolbarItems(lpButtons As TBBUTTON, nButtons As UInteger, uFlags As UInteger) As HRESULT
    End Interface

    <ComImport>
    <Guid("000214E3-0000-0000-C000-000000000046")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Interface IShellView
        Inherits IOleWindow

        Overloads Function GetWindow(<Out> ByRef phwnd As IntPtr) As HRESULT
        Overloads Function ContextSensitiveHelp(fEnterMode As Boolean) As HRESULT
        Function TranslateAccelerator(pmsg As MSG) As HRESULT
        Function EnableModeless(fEnable As Boolean) As HRESULT
        Function UIActivate(uState As UInteger) As HRESULT
        Function Refresh() As HRESULT
        Function CreateViewWindow(psvPrevious As IShellView, pfs As FOLDERSETTINGS, psb As IShellBrowser, prcView As RECT, <Out> ByRef pIntPtr As IntPtr) As HRESULT
        Function DestroyViewWindow() As HRESULT
        Function GetCurrentInfo(<Out> ByRef pfs As FOLDERSETTINGS) As HRESULT
        Function AddPropertySheetPages(dwReserved As Integer, pfn As IntPtr, lparam As IntPtr) As HRESULT
        Function SaveViewState() As HRESULT
        Function SelectItem(pidlItem As IntPtr, uFlags As SVSIF) As HRESULT
        Function GetItemObject(uItem As UInteger, ByRef riid As Guid, <Out> <MarshalAs(UnmanagedType.[Interface])> ByRef ppv As Object) As HRESULT
    End Interface

    <StructLayout(LayoutKind.Sequential)>
    Public Structure MSG
        Public hwnd As IntPtr
        Public message As UInteger
        Public wParam As Integer
        Public lParam As IntPtr
        Public time As Integer
        Public pt As POINT
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure FOLDERSETTINGS
        Public ViewMode As UInteger
        Public fFlags As UInteger
    End Structure

    Public Enum SVSIF
        SVSI_DESELECT = 0
        SVSI_SELECT = &H1
        SVSI_EDIT = &H3
        SVSI_DESELECTOTHERS = &H4
        SVSI_ENSUREVISIBLE = &H8
        SVSI_FOCUSED = &H10
        SVSI_TRANSLATEPT = &H20
        SVSI_SELECTIONMARK = &H40
        SVSI_POSITIONITEM = &H80
        SVSI_CHECK = &H100
        SVSI_CHECK2 = &H200
        SVSI_KEYBOARDSELECT = &H401
        SVSI_NOTAKEFOCUS = &H40000000
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure TBBUTTON
        Public iBitmap As Integer
        Public idCommand As Integer
        Public fsState As Byte
        Public fsStyle As Byte
        Public bReserved0 As Byte
        Public bReserved1 As Byte
        Public dwData As Integer
        Public iString As IntPtr
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure OLEMENUGROUPWIDTHS
        <MarshalAs(UnmanagedType.U2, SizeConst:=6)>
        Public width As Integer()
    End Structure

    <ComImport>
    <Guid("cde725b0-ccc9-4519-917e-325d72fab4ce")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IFolderView
        Function GetCurrentViewMode(ByRef pViewMode As UInteger) As HRESULT
        Function SetCurrentViewMode(ViewMode As UInteger) As HRESULT
        Function GetFolder(ByRef riid As Guid, ByRef ppv As IntPtr) As HRESULT
        Function Item(iItemIndex As Integer, ByRef ppidl As IntPtr) As HRESULT
        Function ItemCount(uFlags As UInteger, ByRef pcItems As Integer) As HRESULT
        Function Items(uFlags As UInteger, ByRef riid As Guid, ByRef ppv As IntPtr) As HRESULT
        Function GetSelectionMarkedItem(ByRef piItem As Integer) As HRESULT
        Function GetFocusedItem(ByRef piItem As Integer) As HRESULT
        Function GetItemPosition(pidl As IntPtr, ByRef ppt As POINT) As HRESULT
        Function GetSpacing(ByRef ppt As POINT) As HRESULT
        Function GetDefaultSpacing(ByRef ppt As POINT) As HRESULT
        Function GetAutoArrange() As HRESULT
        Function SelectItem(iItem As Integer, dwFlags As Integer) As HRESULT
        Function SelectAndPositionItems(cidl As UInteger, apidl As IntPtr, apt As POINT, dwFlags As Integer) As HRESULT
    End Interface

    Public Enum SVGIO
        SVGIO_BACKGROUND = 0
        SVGIO_SELECTION = &H1
        SVGIO_ALLVIEW = &H2
        SVGIO_CHECKED = &H3
        SVGIO_TYPE_MASK = &HF
        SVGIO_FLAG_VIEWORDER = &H80000000
    End Enum

    <ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("000214E6-0000-0000-C000-000000000046")>
    Interface IShellFolder
        Function ParseDisplayName(hwnd As IntPtr, pbc As IntPtr, <MarshalAs(UnmanagedType.LPWStr)> pszDisplayName As String, <[In], Out> ByRef pchEaten As UInteger, <Out> ByRef ppidl As IntPtr, <[In], Out> ByRef pdwAttributes As SFGAO) As HRESULT
        Function EnumObjects(hwnd As IntPtr, grfFlags As SHCONTF, <Out> ByRef ppenumIDList As IEnumIDList) As HRESULT
        Function BindToObject(pidl As IntPtr, pbc As IntPtr, <[In]> ByRef riid As Guid, <Out> <MarshalAs(UnmanagedType.[Interface])> ByRef ppv As Object) As HRESULT
        Function BindToStorage(pidl As IntPtr, pbc As IntPtr, <[In]> ByRef riid As Guid, <Out> <MarshalAs(UnmanagedType.[Interface])> ByRef ppv As Object) As HRESULT
        Function CompareIDs(lParam As IntPtr, pidl1 As IntPtr, pidl2 As IntPtr) As HRESULT
        Function CreateViewObject(hwndOwner As IntPtr, <[In]> ByRef riid As Guid, <Out> <MarshalAs(UnmanagedType.[Interface])> ByRef ppv As Object) As HRESULT
        Function GetAttributesOf(cidl As UInteger, apidl As IntPtr, <[In], Out> ByRef rgfInOut As SFGAO) As HRESULT
        Function GetUIObjectOf(hwndOwner As IntPtr, cidl As UInteger, ByRef apidl As IntPtr, <[In]> ByRef riid As Guid, <[In], Out> ByRef rgfReserved As UInteger, <Out> ByRef ppv As IntPtr) As HRESULT
        Function GetDisplayNameOf(pidl As IntPtr, uFlags As SHGDNF, <Out> ByRef pName As STRRET) As HRESULT
        Function SetNameOf(hwnd As IntPtr, pidl As IntPtr, <MarshalAs(UnmanagedType.LPWStr)> pszName As String, uFlags As SHGDNF, <Out> ByRef ppidlOut As IntPtr) As HRESULT
    End Interface

    Public Enum SHCONTF
        SHCONTF_CHECKING_FOR_CHILDREN = &H10
        SHCONTF_FOLDERS = &H20
        SHCONTF_NONFOLDERS = &H40
        SHCONTF_INCLUDEHIDDEN = &H80
        SHCONTF_INIT_ON_FIRST_NEXT = &H100
        SHCONTF_NETPRINTERSRCH = &H200
        SHCONTF_SHAREABLE = &H400
        SHCONTF_STORAGE = &H800
        SHCONTF_NAVIGATION_ENUM = &H1000
        SHCONTF_FASTITEMS = &H2000
        SHCONTF_FLATLIST = &H4000
        SHCONTF_ENABLE_ASYNC = &H8000
    End Enum

    Public Enum SFGAO
        CANCOPY = &H1
        CANMOVE = &H2
        CANLINK = &H4
        STORAGE = &H8
        CANRENAME = &H10
        CANDELETE = &H20
        HASPROPSHEET = &H40
        DROPTARGET = &H100
        CAPABILITYMASK = &H177
        ENCRYPTED = &H2000
        ISSLOW = &H4000
        GHOSTED = &H8000
        LINK = &H10000
        SHARE = &H20000
        [READONLY] = &H40000
        HIDDEN = &H80000
        DISPLAYATTRMASK = &HFC000
        STREAM = &H400000
        STORAGEANCESTOR = &H800000
        VALIDATE = &H1000000
        REMOVABLE = &H2000000
        COMPRESSED = &H4000000
        BROWSABLE = &H8000000
        FILESYSANCESTOR = &H10000000
        FOLDER = &H20000000
        FILESYSTEM = &H40000000
        HASSUBFOLDER = &H80000000
        CONTENTSMASK = &H80000000
        STORAGECAPMASK = &H70C50008
        PKEYSFGAOMASK = &H81044000
    End Enum

    Public Enum SHGDNF
        SHGDN_NORMAL = 0
        SHGDN_INFOLDER = &H1
        SHGDN_FOREDITING = &H1000
        SHGDN_FORADDRESSBAR = &H4000
        SHGDN_FORPARSING = &H8000
    End Enum

    <StructLayout(LayoutKind.Explicit, Size:=264)>
    Public Structure STRRET
        <FieldOffset(0)>
        Public uType As UInteger
        <FieldOffset(4)>
        Public pOleStr As IntPtr
        <FieldOffset(4)>
        Public uOffset As UInteger
        <FieldOffset(4)>
        Public cString As IntPtr
    End Structure

    <ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("000214F2-0000-0000-C000-000000000046")>
    Interface IEnumIDList
        <PreserveSig()>
        Function [Next](celt As UInteger, <Out> ByRef rgelt As IntPtr, <Out> ByRef pceltFetched As Integer) As HRESULT
        <PreserveSig()>
        Function Skip(celt As UInteger) As HRESULT
        Sub Reset()
        Function Clone() As IEnumIDList
    End Interface

    <DllImport("Shlwapi.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Public Function StrRetToBuf(ByRef pstr As STRRET, ByVal pidl As IntPtr, ByVal pszBuf As System.Text.StringBuilder, <MarshalAs(UnmanagedType.U4)> ByVal cchBuf As UInteger) As HRESULT
    End Function
    Sub Main()
        Dim pShellWindows = New SHDocVw.ShellWindows()
        Dim hWndDesktop As IntPtr
        Dim oDispatch = pShellWindows.FindWindowSW(0, Nothing, SWC_DESKTOP, hWndDesktop, ShellWindowFindWindowOptions.SWFO_NEEDDISPATCH)
        Dim pDispatch = Marshal.GetIDispatchForObject(oDispatch)
        Dim pvOut As IntPtr = IntPtr.Zero
        Dim SID_STopLevelBrowser As Guid = New Guid("4C96BE40-915C-11CF-99D3-00AA004AE837")
        Dim IID_IShellBrowser As Guid = New Guid("000214E2-0000-0000-C000-000000000046")
        IUnknown_QueryService(pDispatch, SID_STopLevelBrowser, IID_IShellBrowser, pvOut)
        If (pvOut <> IntPtr.Zero) Then
            Dim psb As IShellBrowser = TryCast(Marshal.GetObjectForIUnknown(pvOut), IShellBrowser)
            If (psb IsNot Nothing) Then
                Dim psv As IShellView = Nothing
                Dim hr As HRESULT = psb.QueryActiveShellView(psv)
                If (hr = HRESULT.S_OK) Then
                    Dim pFolderView = TryCast(psv, IFolderView)
                    Dim pShellFolderPtr As IntPtr = IntPtr.Zero
                    hr = pFolderView.GetFolder(GetType(IShellFolder).GUID, pShellFolderPtr)
                    Dim psf As IShellFolder = TryCast(Marshal.GetObjectForIUnknown(pShellFolderPtr), IShellFolder)
                    Dim nItemCount As Integer = 0
                    hr = pFolderView.ItemCount(SVGIO.SVGIO_ALLVIEW, nItemCount)
                    If (hr = HRESULT.S_OK) Then
                        For i As Integer = 0 To nItemCount - 1
                            Dim pidlItem As IntPtr
                            hr = pFolderView.Item(i, pidlItem)
                            If (hr = HRESULT.S_OK) Then
                                Console.WriteLine(String.Format("Item {0}", (i + 1).ToString()))
                                Dim strretFolderName As STRRET
                                hr = psf.GetDisplayNameOf(pidlItem, SHGDNF.SHGDN_NORMAL, strretFolderName)
                                Dim sDisplayName As String = Nothing
                                If (hr = HRESULT.S_OK) Then
                                    Dim sbDisplayName As System.Text.StringBuilder
                                    sbDisplayName = New System.Text.StringBuilder(256)
                                    StrRetToBuf(strretFolderName, pidlItem, sbDisplayName, sbDisplayName.Capacity)
                                    sDisplayName = sbDisplayName.ToString
                                End If
                                Console.WriteLine(String.Format("{0}Name : {1}", vbTab, sDisplayName))
                            End If
                            Dim nPos As POINT = Nothing
                            hr = pFolderView.GetItemPosition(pidlItem, nPos)
                            Console.WriteLine(String.Format("{0}Position : ({1},{2})", vbTab, nPos.X, nPos.Y))
                        Next
                    End If
                    Marshal.ReleaseComObject(psf)
                    Marshal.ReleaseComObject(psv)
                End If
                Marshal.ReleaseComObject(psb)
            End If
        End If
    End Sub

End Module
