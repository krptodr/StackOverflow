Imports IDataObject_Com = System.Runtime.InteropServices.ComTypes.IDataObject
Imports System.Windows.Interop
Imports System.Runtime.InteropServices
Imports WindowsApplication2

Public Class DragAndDrop
    Implements IOleDropTarget

    Public Function OleDragEnter(<[In]> <MarshalAs(UnmanagedType.Interface)> pDataObj As Object, <[In]> <MarshalAs(UnmanagedType.U4)> grfKeyState As Integer, <[In]> <MarshalAs(UnmanagedType.U8)> pt As Long, <[In]> <Out> ByRef pdwEffect As Integer) As Integer Implements IOleDropTarget.OleDragEnter
        Debug.WriteLine("New Drag Enter message")
        Return NativeMethodsEx.S_OK
    End Function

    Public Function OleDragLeave() As Integer Implements IOleDropTarget.OleDragLeave
        Throw New NotImplementedException()
    End Function

    Public Function OleDragOver(<[In]> <MarshalAs(UnmanagedType.U4)> grfKeyState As Integer, <[In]> <MarshalAs(UnmanagedType.U8)> pt As Long, <[In]> <Out> ByRef pdwEffect As Integer) As Integer Implements IOleDropTarget.OleDragOver
        Throw New NotImplementedException()
    End Function

    Public Function OleDrop(<[In]> <MarshalAs(UnmanagedType.Interface)> pDataObj As Object, <[In]> <MarshalAs(UnmanagedType.U4)> grfKeyState As Integer, <[In]> <MarshalAs(UnmanagedType.U8)> pt As Long, <[In]> <Out> ByRef pdwEffect As Integer) As Integer Implements IOleDropTarget.OleDrop
        Throw New NotImplementedException()
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
