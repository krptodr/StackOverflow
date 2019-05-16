Imports IDataObject_Com = System.Runtime.InteropServices.ComTypes.IDataObject
Imports System.Windows.Interop
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Runtime.InteropServices.ComTypes

Namespace Browser
    <StructLayout(LayoutKind.Sequential)>
    Public Structure Win32Point
        Public x As Integer
        Public y As Integer
    End Structure

    Module Win32API
        Public Const Chrome_WidgetWin As String = "Chrome_WidgetWin_0"
        Public outparam As IntPtr
        Public Function TryFindHandl(ByVal browserHandle As IntPtr, <Out> ByRef chromeWidgetHostHandle As IntPtr) As Boolean

            Dim cbXL As New NativeMethodsEx.EnumChildCallback(AddressOf EnumChildProc_Browser)
            NativeMethodsEx.EnumChildWindows(browserHandle, cbXL, chromeWidgetHostHandle)

            Return chromeWidgetHostHandle <> IntPtr.Zero

        End Function

        Public Function EnumChildProc_Browser(ByVal hwndChild As Integer, ByRef lParam As Integer) As Boolean
            Dim buf As New StringBuilder(128)
            'Dim hwndMyWindow As IntPtr = FindWindowEx(NULL, NULL, "mywindowclass", NULL);
            NativeMethodsEx.GetClassName(hwndChild, buf, 128)
            Dim ret = NativeMethodsEx.RevokeDragDrop(hwndChild)

            If ret = NativeMethodsEx.DRAGDROP_E_NOTREGISTERED Then
                Debug.WriteLine("No Drag and Drop Registered")
            End If

            If buf.ToString = Chrome_WidgetWin Then
                lParam = hwndChild
                Return False
            End If
            Return True
        End Function
    End Module


    <ComImport>
    <Guid("4657278A-411B-11d2-839A-00C04FD918D0")>
    Public Class DragDropHelper
        Implements IDropTargetHelper


        Public Sub DragEnter(<[In]> hwndTarget As IntPtr, <[In]> <MarshalAs(UnmanagedType.Interface)> dataObject As IDataObject_Com, <[In]> ByRef pt As Win32Point, <[In]> effect As Integer) Implements IDropTargetHelper.DragEnter
            Throw New NotImplementedException()
        End Sub

        Public Sub DragLeave() Implements IDropTargetHelper.DragLeave
            Throw New NotImplementedException()
        End Sub

        Public Sub DragOver(<[In]> ByRef pt As Win32Point, <[In]> effect As Integer) Implements IDropTargetHelper.DragOver
            Throw New NotImplementedException()
        End Sub

        Public Sub Drop(<[In]> <MarshalAs(UnmanagedType.Interface)> dataObject As IDataObject_Com, <[In]> ByRef pt As Win32Point, <[In]> effect As Integer) Implements IDropTargetHelper.Drop
            Throw New NotImplementedException()
        End Sub

        Public Sub Show(<[In]> show As Boolean) Implements IDropTargetHelper.Show
            Throw New NotImplementedException()
        End Sub
    End Class

    <ComVisible(True)>
    <ComImport>
    <Guid("4657278B-411B-11D2-839A-00C04FD918D0")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Interface IDropTargetHelper
        Sub DragEnter(
 <[In]> ByVal hwndTarget As IntPtr,
 <[In], MarshalAs(UnmanagedType.[Interface])> ByVal dataObject As IDataObject_Com,
 <[In]> ByRef pt As Win32Point,
 <[In]> ByVal effect As Integer)
        Sub DragLeave()
        Sub DragOver(
 <[In]> ByRef pt As Win32Point,
 <[In]> ByVal effect As Integer)
        Sub Drop(
 <[In], MarshalAs(UnmanagedType.[Interface])> ByVal dataObject As IDataObject_Com,
 <[In]> ByRef pt As Win32Point,
 <[In]> ByVal effect As Integer)
        Sub Show(
 <[In]> ByVal show As Boolean)
    End Interface



End Namespace