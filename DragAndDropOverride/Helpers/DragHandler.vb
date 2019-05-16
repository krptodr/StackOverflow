

Imports CefSharp
Imports CefSharp.Enums

Public Class DragHandler
    Implements IDragHandler

    Public Sub OnDraggableRegionsChanged(chromiumWebBrowser As IWebBrowser, browser As IBrowser, regions As IList(Of DraggableRegion)) Implements IDragHandler.OnDraggableRegionsChanged
        Throw New NotImplementedException()
    End Sub

    Public Function OnDragEnter(chromiumWebBrowser As IWebBrowser, browser As IBrowser, dragData As IDragData, mask As DragOperationsMask) As Boolean Implements IDragHandler.OnDragEnter
        ' MsgBox("Test")
        Dim DataObject As OutlookDataObject = New OutlookDataObject(CType(CType(dragData, IDataObject), System.Windows.Forms.IDataObject))
        Return False
    End Function
End Class
