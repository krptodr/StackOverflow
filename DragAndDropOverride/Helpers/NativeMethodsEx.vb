Imports DragAndDropOverride.Browser
Imports System.Runtime.InteropServices

'Public NotInheritable Class NativeMethodsEx

'End Class



Partial Public Class NativeMethodsEx
    Friend Const DRAGDROP_E_NOTREGISTERED = &H80040100
    Friend Const DRAGDROP_E_INVALIDHWND = &H80040102
    Friend Const DRAGDROP_E_ALREADYREGISTERED = &H80040101
    Friend Const E_OUTOFMEMORY = &H8007000E
    Friend Const S_OK = 0


    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Friend Shared Function GetClassName(ByVal hWnd As System.IntPtr, ByVal lpClassname As System.Text.StringBuilder, ByVal nMaxCount As Integer) As Integer
    End Function

    Friend Delegate Function EnumChildCallback(ByVal hwnd As Integer, ByRef lParam As Integer) As Boolean

    '<DllImport("User32.dll")>
    'Friend Shared Function EnumWindows(ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

    'End Function


    <DllImport("User32.dll")>
    Friend Shared Function EnumChildWindows(ByVal hWndParent As Integer, ByVal lpEnumFunc As EnumChildCallback, ByRef lParam As Integer) As Boolean
    End Function

    <DllImport("ole32.dll")>
    Friend Shared Function RegisterDragDrop(ByVal hwnd As IntPtr, DropTarget As IOleDropTarget) As IntPtr
    End Function

    <DllImport("ole32.dll")>
    Friend Shared Function RevokeDragDrop(ByVal hwnd As IntPtr) As IntPtr
    End Function

    '''Return Type: BOOL->int
    '''param0: HWND->HWND__*
    '''param1: LPARAM->LONG_PTR->int
    <System.Runtime.InteropServices.UnmanagedFunctionPointerAttribute(System.Runtime.InteropServices.CallingConvention.StdCall)>
    Public Delegate Function WNDENUMPROC(ByVal param0 As System.IntPtr, ByVal param1 As System.IntPtr) As Integer


    '''Return Type: BOOL->int
    '''lpEnumFunc: WNDENUMPROC
    '''lParam: LPARAM->LONG_PTR->int
    <System.Runtime.InteropServices.DllImportAttribute("user32.dll", EntryPoint:="EnumWindows")>
    Public Shared Function EnumWindows(ByVal lpEnumFunc As WNDENUMPROC, <System.Runtime.InteropServices.MarshalAsAttribute(System.Runtime.InteropServices.UnmanagedType.SysInt)> ByVal lParam As Integer) As <System.Runtime.InteropServices.MarshalAsAttribute(System.Runtime.InteropServices.UnmanagedType.Bool)> Boolean
    End Function
End Class