Imports IDataObject_Com = System.Runtime.InteropServices.ComTypes.IDataObject
Imports System.Windows.Interop
Imports System.Runtime.InteropServices
Imports DragAndDropOverride.Browser
Imports DragAndDropOverride
Imports CefSharp.WinForms

Public Class DragAndDrop
    Implements IOleDropTarget

    Public Function OleDragEnter(<[In]> <MarshalAs(UnmanagedType.Interface)> pDataObj As Object, <[In]> <MarshalAs(UnmanagedType.U4)> grfKeyState As Integer, <[In]> <MarshalAs(UnmanagedType.U8)> pt As Long, <[In]> <Out> ByRef pdwEffect As Integer) As Integer Implements IOleDropTarget.OleDragEnter
        Debug.WriteLine("New Drag Enter message")
        Return NativeMethodsEx.S_OK
    End Function

    Public Function OleDragLeave() As Integer Implements IOleDropTarget.OleDragLeave
        Return NativeMethodsEx.S_OK
    End Function

    Public Function OleDragOver(<[In]> <MarshalAs(UnmanagedType.U4)> grfKeyState As Integer, <[In]> <MarshalAs(UnmanagedType.U8)> pt As Long, <[In]> <Out> ByRef pdwEffect As Integer) As Integer Implements IOleDropTarget.OleDragOver
        Return NativeMethodsEx.S_OK
    End Function

    Public Function OleDrop(<[In]> <MarshalAs(UnmanagedType.Interface)> pDataObj As Object, <[In]> <MarshalAs(UnmanagedType.U4)> grfKeyState As Integer, <[In]> <MarshalAs(UnmanagedType.U8)> pt As Long, <[In]> <Out> ByRef pdwEffect As Integer) As Integer Implements IOleDropTarget.OleDrop
        Debug.WriteLine("New Drag Enter message")
        Dim winPT As Win32Point
        winPT.x = CInt(pt And &H7FFFFFFF)
        winPT.y = CInt((pt >> 32) And &H7FFFFFFF)
        Dim eff As DragDropEffects = DragDropEffects.None
        'this is my event I am sending back to the browser class to deal with.
        'RaiseEvent DBDragEnter(eff, New Point(winPT.x, winPT.y))
        'you need to pass in the effect
        pdwEffect = CInt(eff)
        'this is the helper which shows the nice icon you drag around.
        Dim ddHelper As New DragDropHelper
        ddHelper.DragEnter(outparam, CType(pDataObj, IDataObject_Com), winPT, CInt(eff))
        Return NativeMethodsEx.S_OK
        'Return NativeMethodsEx.S_OK
    End Function
End Class
<ComImport, Guid("00000122-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Interface IOleDropTarget
    <PreserveSig>
    Function OleDragEnter(
    <[In], MarshalAs(UnmanagedType.[Interface])> ByVal pDataObj As Object,
    <[In], MarshalAs(UnmanagedType.U4)> ByVal grfKeyState As Integer,
    <[In], MarshalAs(UnmanagedType.U8)> ByVal pt As Long,
    <[In], Out> ByRef pdwEffect As Integer) As Integer
    <PreserveSig>
    Function OleDragOver(
    <[In], MarshalAs(UnmanagedType.U4)> ByVal grfKeyState As Integer,
    <[In], MarshalAs(UnmanagedType.U8)> ByVal pt As Long,
    <[In], Out> ByRef pdwEffect As Integer) As Integer
    <PreserveSig>
    Function OleDragLeave() As Integer
    <PreserveSig>
    Function OleDrop(
    <[In], MarshalAs(UnmanagedType.[Interface])> ByVal pDataObj As Object,
    <[In], MarshalAs(UnmanagedType.U4)> ByVal grfKeyState As Integer,
    <[In], MarshalAs(UnmanagedType.U8)> ByVal pt As Long,
    <[In], Out> ByRef pdwEffect As Integer) As Integer
End Interface
