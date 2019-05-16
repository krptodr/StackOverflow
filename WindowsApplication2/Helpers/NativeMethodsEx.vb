Imports WindowsApplication2.Browser
Imports System.Runtime.InteropServices

Public NotInheritable Class NativeMethodsEx
        Friend Const DRAGDROP_E_NOTREGISTERED = &H80040100
        Friend Const DRAGDROP_E_INVALIDHWND = &H80040102
        Friend Const DRAGDROP_E_ALREADYREGISTERED = &H80040101
        Friend Const E_OUTOFMEMORY = &H8007000E
        Friend Const S_OK = 0


        <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
        Friend Shared Function GetClassName(ByVal hWnd As System.IntPtr, ByVal lpClassname As System.Text.StringBuilder, ByVal nMaxCount As Integer) As Integer
        End Function

        Friend Delegate Function EnumChildCallback(ByVal hwnd As Integer, ByRef lParam As Integer) As Boolean

        <DllImport("User32.dll")>
        Friend Shared Function EnumChildWindows(ByVal hWndParent As Integer, ByVal lpEnumFunc As EnumChildCallback, ByRef lParam As Integer) As Boolean
        End Function


    <DllImport("ole32.dll")>
    Friend Shared Function RegisterDragDrop(ByVal hwnd As IntPtr, DropTarget As IOleDropTarget) As IntPtr
    End Function

    <DllImport("ole32.dll")>
        Friend Shared Function RevokeDragDrop(ByVal hwnd As IntPtr) As IntPtr
        End Function
End Class