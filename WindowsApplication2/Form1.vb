Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports System.Text
Imports CefSharp
Imports CefSharp.Enums
Imports CefSharp.WinForms
Imports CefSharp.WinForms.Internals
Imports WindowsApplication2
Imports WindowsApplication2.Browser
Imports WindowsApplication2.iwantedue.Windows.Forms

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

Public Class Form1


    Public Property MBrowser As IWinFormsWebBrowser

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        'Dim browser = New ChromiumWebBrowser("Google.com") With {.Dock = DockStyle.Fill}
        'pBrowser.Controls.Add(browser)
        ChromiumWebBrowser1.Load("google.com")

        'ChromiumWebBrowser1.AllowDrop = True

        MBrowser = ChromiumWebBrowser1
        'Self Register Drag and Drop
        'Dim drag As New DragDropHelper
        Dim outparam As IntPtr
        If DragDropHelper.TryFindHandl(ChromiumWebBrowser1.Handle, outparam) Then
            Debug.WriteLine("Handle Removed")
            NativeMethodsEx.RegisterDragDrop(Me.Handle, New DragAndDrop)
        End If
        'browser.AllowDrop = True
        'MBrowser.DragHandler = New DragHandler

        AddHandler MBrowser.AddressChanged, AddressOf OnBrowserAddressChanged

        ' Add any initialization after the InitializeComponent() call.

    End Sub



    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load









    End Sub

    Private Sub Form1_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs)
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        ElseIf e.Data.GetDataPresent("FileGroupDescriptor") Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub
    Private Sub Form1_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs)
        Dim fileNames As String() = Nothing

        Try

            If e.Data.GetDataPresent(DataFormats.FileDrop, False) = True Then
                fileNames = CType(e.Data.GetData(DataFormats.FileDrop), String())

                For Each fileName As String In fileNames
                Next
            ElseIf e.Data.GetDataPresent("FileGroupDescriptor") Then
                Dim theStream As Stream = CType(e.Data.GetData("FileGroupDescriptor"), Stream)
                Dim fileGroupDescriptor As Byte() = New Byte(511) {}
                theStream.Read(fileGroupDescriptor, 0, 512)
                Dim fileName As StringBuilder = New StringBuilder("")
                Dim i As Integer = 76

                While fileGroupDescriptor(i) <> 0
                    fileName.Append(Convert.ToChar(fileGroupDescriptor(i)))
                    i += 1
                End While

                theStream.Close()
                Dim fpath As String = Path.GetTempPath()
                Dim theFile As String = fpath & fileName.ToString()
                Dim ms As MemoryStream = CType(e.Data.GetData("FileContents", True), MemoryStream)
                Dim fileBytes As Byte() = New Byte(ms.Length - 1) {}
                ms.Position = 0
                ms.Read(fileBytes, 0, CInt(ms.Length))
                Dim fs As FileStream = New FileStream(theFile, FileMode.Create)
                fs.Write(fileBytes, 0, CInt(fileBytes.Length))
                fs.Close()
                Dim tempFile As FileInfo = New FileInfo(theFile)

                If tempFile.Exists = True Then
                    tempFile.Delete()
                Else
                    Trace.WriteLine("File was not created!")
                End If
            End If

        Catch ex As Exception
            Trace.WriteLine("Error in DragDrop function: " & ex.Message)
        End Try
    End Sub

    Public Sub OnBrowserAddressChanged(sender As Object, args As AddressChangedEventArgs)
        Me.InvokeOnUiThreadIfRequired(Sub()
                                          TextBox1.Text = args.Address
                                      End Sub)
        'TextBox1.Text = args.Address

    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub pBrowser_Paint(sender As Object, e As PaintEventArgs) Handles pBrowser.Paint

    End Sub

    Private Sub TextBox1_KeyUp(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyUp
        If Not e.KeyCode = Keys.Enter Then
            Exit Sub
        End If
        MBrowser.Load(TextBox1.Text)
    End Sub



End Class
