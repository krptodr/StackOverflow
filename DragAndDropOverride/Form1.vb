Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports System.Text
Imports CefSharp
Imports CefSharp.Enums
Imports CefSharp.WinForms
Imports CefSharp.WinForms.Internals
Imports DragAndDropOverride
Imports DragAndDropOverride.Browser


Public Class Form1

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click

        If Win32API.TryFindHandl(ChromiumWebBrowser1.Handle, outparam) Then
            Debug.WriteLine("Handle Removed")
            NativeMethodsEx.RegisterDragDrop(outparam, New DragAndDrop)
        End If
        Exit Sub
        ' clear list
        hWndList.Clear()
        ' enum windows, find classname "CabinetWClass", add handles to list
        EnumWindows(AddressOf ClassesByName, IntPtr.Zero)
        ' display list contents (window handles).
        If hWndList.Count = 0 Then
            MessageBox.Show("None found, list is empty!")
        Else
            For Each hWnd In hWndList
                Debug.WriteLine(hWnd)
            Next
        End If
    End Sub

    Private hWndList As New List(Of IntPtr)
    ' API Stuff...
    Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As IntPtr, ByVal lpClassName As StringBuilder, ByVal nMaxCount As Int32) As Int32
    Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As EnumWindowsProcDelegate, ByVal lParam As IntPtr) As Boolean
    Private Delegate Function EnumWindowsProcDelegate(ByVal hWnd As IntPtr, ByVal lParam As IntPtr) As Boolean

    Private Function ClassesByName(ByVal hWnd As IntPtr, ByVal lParam As IntPtr) As Boolean
        If Get_ClassName(hWnd) = Win32API.Chrome_WidgetWin Then ' "CabinetWClass", used in Win7 Explorer windows.
            hWndList.Add(hWnd)
        End If
        Return True
    End Function

    Private Function Get_ClassName(ByVal hWnd As IntPtr) As String
        Dim sb As New StringBuilder(256)
        GetClassName(hWnd, sb, 256)
        Return sb.ToString
    End Function

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
