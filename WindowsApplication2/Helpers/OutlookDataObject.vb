Imports System
Imports System.IO
Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports System.Text
Imports System.Reflection
Imports System.Windows.Forms

Namespace iwantedue.Windows.Forms
    Public Class OutlookDataObject
        Implements System.Windows.Forms.IDataObject

        Private Class NativeMethods
            <DllImport("kernel32.dll")>
            Private Shared Function GlobalLock(ByVal hMem As IntPtr) As IntPtr

            End Function
            <DllImport("ole32.dll", PreserveSig:=False)>
            Public Shared Function CreateILockBytesOnHGlobal(ByVal hGlobal As IntPtr, ByVal fDeleteOnRelease As Boolean) As ILockBytes

            End Function
            <DllImport("OLE32.DLL", CharSet:=CharSet.Auto, PreserveSig:=False)>
            Public Shared Function GetHGlobalFromILockBytes(ByVal pLockBytes As ILockBytes) As IntPtr

            End Function
            <DllImport("OLE32.DLL", CharSet:=CharSet.Unicode, PreserveSig:=False)>
            Public Shared Function StgCreateDocfileOnILockBytes(ByVal plkbyt As ILockBytes, ByVal grfMode As UInteger, ByVal reserved As UInteger) As IStorage

            End Function

            <ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("0000000B-0000-0000-C000-000000000046")>
            Interface IStorage

                Function CreateStream(
                <[In], MarshalAs(UnmanagedType.BStr)> ByVal pwcsName As String,
                <[In], MarshalAs(UnmanagedType.U4)> ByVal grfMode As Integer,
                <[In], MarshalAs(UnmanagedType.U4)> ByVal reserved1 As Integer,
                <[In], MarshalAs(UnmanagedType.U4)> ByVal reserved2 As Integer) As IStream

                Function OpenStream(
                <[In], MarshalAs(UnmanagedType.BStr)> ByVal pwcsName As String, ByVal reserved1 As IntPtr,
                <[In], MarshalAs(UnmanagedType.U4)> ByVal grfMode As Integer,
                <[In], MarshalAs(UnmanagedType.U4)> ByVal reserved2 As Integer) As IStream

                Function CreateStorage(
                <[In], MarshalAs(UnmanagedType.BStr)> ByVal pwcsName As String,
                <[In], MarshalAs(UnmanagedType.U4)> ByVal grfMode As Integer,
                <[In], MarshalAs(UnmanagedType.U4)> ByVal reserved1 As Integer,
                <[In], MarshalAs(UnmanagedType.U4)> ByVal reserved2 As Integer) As IStorage

                Function OpenStorage(
                <[In], MarshalAs(UnmanagedType.BStr)> ByVal pwcsName As String, ByVal pstgPriority As IntPtr,
                <[In], MarshalAs(UnmanagedType.U4)> ByVal grfMode As Integer, ByVal snbExclude As IntPtr,
                <[In], MarshalAs(UnmanagedType.U4)> ByVal reserved As Integer) As IStorage
                Sub CopyTo(ByVal ciidExclude As Integer,
<[In], MarshalAs(UnmanagedType.LPArray)> ByVal pIIDExclude As Guid(), ByVal snbExclude As IntPtr,
<[In], MarshalAs(UnmanagedType.Interface)> ByVal stgDest As IStorage)
                Sub MoveElementTo(
<[In], MarshalAs(UnmanagedType.BStr)> ByVal pwcsName As String,
<[In], MarshalAs(UnmanagedType.Interface)> ByVal stgDest As IStorage,
<[In], MarshalAs(UnmanagedType.BStr)> ByVal pwcsNewName As String,
<[In], MarshalAs(UnmanagedType.U4)> ByVal grfFlags As Integer)
                Sub Commit(ByVal grfCommitFlags As Integer)
                Sub Revert()
                Sub EnumElements(
<[In], MarshalAs(UnmanagedType.U4)> ByVal reserved1 As Integer, ByVal reserved2 As IntPtr,
<[In], MarshalAs(UnmanagedType.U4)> ByVal reserved3 As Integer, <Out>
<MarshalAs(UnmanagedType.Interface)> ByRef ppVal As Object)
                Sub DestroyElement(
<[In], MarshalAs(UnmanagedType.BStr)> ByVal pwcsName As String)
                Sub RenameElement(
<[In], MarshalAs(UnmanagedType.BStr)> ByVal pwcsOldName As String,
<[In], MarshalAs(UnmanagedType.BStr)> ByVal pwcsNewName As String)
                Sub SetElementTimes(
<[In], MarshalAs(UnmanagedType.BStr)> ByVal pwcsName As String,
<[In]> ByVal pctime As System.Runtime.InteropServices.ComTypes.FILETIME,
<[In]> ByVal patime As System.Runtime.InteropServices.ComTypes.FILETIME,
<[In]> ByVal pmtime As System.Runtime.InteropServices.ComTypes.FILETIME)
                Sub SetClass(
<[In]> ByRef clsid As Guid)
                Sub SetStateBits(ByVal grfStateBits As Integer, ByVal grfMask As Integer)
                Sub Stat(
<Out> ByRef pStatStg As System.Runtime.InteropServices.ComTypes.STATSTG, ByVal grfStatFlag As Integer)
            End Interface

            <ComImport, Guid("0000000A-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
            Interface ILockBytes
                Sub ReadAt(
<[In], MarshalAs(UnmanagedType.U8)> ByVal ulOffset As Long,
<Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> ByVal pv As Byte(),
<[In], MarshalAs(UnmanagedType.U4)> ByVal cb As Integer,
<Out, MarshalAs(UnmanagedType.LPArray)> ByVal pcbRead As Integer())
                Sub WriteAt(
<[In], MarshalAs(UnmanagedType.U8)> ByVal ulOffset As Long, ByVal pv As IntPtr,
<[In], MarshalAs(UnmanagedType.U4)> ByVal cb As Integer,
<Out, MarshalAs(UnmanagedType.LPArray)> ByVal pcbWritten As Integer())
                Sub Flush()
                Sub SetSize(
<[In], MarshalAs(UnmanagedType.U8)> ByVal cb As Long)
                Sub LockRegion(
<[In], MarshalAs(UnmanagedType.U8)> ByVal libOffset As Long,
<[In], MarshalAs(UnmanagedType.U8)> ByVal cb As Long,
<[In], MarshalAs(UnmanagedType.U4)> ByVal dwLockType As Integer)
                Sub UnlockRegion(
<[In], MarshalAs(UnmanagedType.U8)> ByVal libOffset As Long,
<[In], MarshalAs(UnmanagedType.U8)> ByVal cb As Long,
<[In], MarshalAs(UnmanagedType.U4)> ByVal dwLockType As Integer)
                Sub Stat(<Out> ByRef pstatstg As System.Runtime.InteropServices.ComTypes.STATSTG,
<[In], MarshalAs(UnmanagedType.U4)> ByVal grfStatFlag As Integer)
            End Interface

            <StructLayout(LayoutKind.Sequential)>
            Public NotInheritable Class POINTL
                Public x As Integer
                Public y As Integer
            End Class

            <StructLayout(LayoutKind.Sequential)>
            Public NotInheritable Class SIZEL
                Public cx As Integer
                Public cy As Integer
            End Class

            <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Ansi)>
            Public NotInheritable Class FILEGROUPDESCRIPTORA
                Public cItems As UInteger
                Public fgd As FILEDESCRIPTORA()
            End Class

            <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Ansi)>
            Public NotInheritable Class FILEDESCRIPTORA
                Public dwFlags As UInteger
                Public clsid As Guid
                Public sizel As SIZEL
                Public pointl As POINTL
                Public dwFileAttributes As UInteger
                Public ftCreationTime As System.Runtime.InteropServices.ComTypes.FILETIME
                Public ftLastAccessTime As System.Runtime.InteropServices.ComTypes.FILETIME
                Public ftLastWriteTime As System.Runtime.InteropServices.ComTypes.FILETIME
                Public nFileSizeHigh As UInteger
                Public nFileSizeLow As UInteger
                <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)>
                Public cFileName As String
            End Class

            <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
            Public NotInheritable Class FILEGROUPDESCRIPTORW
                Public cItems As UInteger
                Public fgd As FILEDESCRIPTORW()
            End Class

            <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
            Public NotInheritable Class FILEDESCRIPTORW
                Public dwFlags As UInteger
                Public clsid As Guid
                Public sizel As SIZEL
                Public pointl As POINTL
                Public dwFileAttributes As UInteger
                Public ftCreationTime As System.Runtime.InteropServices.ComTypes.FILETIME
                Public ftLastAccessTime As System.Runtime.InteropServices.ComTypes.FILETIME
                Public ftLastWriteTime As System.Runtime.InteropServices.ComTypes.FILETIME
                Public nFileSizeHigh As UInteger
                Public nFileSizeLow As UInteger
                <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)>
                Public cFileName As String
            End Class
        End Class

        Private underlyingDataObject As System.Windows.Forms.IDataObject
        Private comUnderlyingDataObject As System.Runtime.InteropServices.ComTypes.IDataObject
        Private oleUnderlyingDataObject As System.Windows.Forms.IDataObject
        Private getDataFromHGLOBLALMethod As MethodInfo

        Public Sub New(ByVal underlyingDataObject As System.Windows.Forms.IDataObject)
            Me.underlyingDataObject = underlyingDataObject
            Me.comUnderlyingDataObject = CType(Me.underlyingDataObject, System.Runtime.InteropServices.ComTypes.IDataObject)
            Dim innerDataField As FieldInfo = Me.underlyingDataObject.[GetType]().GetField("innerData", BindingFlags.NonPublic Or BindingFlags.Instance)
            Me.oleUnderlyingDataObject = CType(innerDataField.GetValue(Me.underlyingDataObject), System.Windows.Forms.IDataObject)
            Me.getDataFromHGLOBLALMethod = Me.oleUnderlyingDataObject.[GetType]().GetMethod("GetDataFromHGLOBLAL", BindingFlags.NonPublic Or BindingFlags.Instance)
        End Sub

#Region "IDataObject Members"

        Public Function GetData(ByVal format As Type) As Object Implements System.Windows.Forms.IDataObject.GetData

            Return Me.GetData(format.FullName)
        End Function

        Public Function GetData(ByVal format As String) As Object Implements System.Windows.Forms.IDataObject.GetData
            Return Me.GetData(format, True)
        End Function

        Public Function GetData(ByVal format As String, ByVal autoConvert As Boolean) As Object Implements System.Windows.Forms.IDataObject.GetData
            Select Case format
                Case "FileGroupDescriptor"
                    Dim fileGroupDescriptorAPointer As IntPtr = IntPtr.Zero

                    Try
                        Dim fileGroupDescriptorStream As MemoryStream = CType(Me.underlyingDataObject.GetData("FileGroupDescriptor", autoConvert), MemoryStream)
                        Dim fileGroupDescriptorBytes As Byte() = New Byte(fileGroupDescriptorStream.Length - 1) {}
                        fileGroupDescriptorStream.Read(fileGroupDescriptorBytes, 0, fileGroupDescriptorBytes.Length)
                        fileGroupDescriptorStream.Close()
                        fileGroupDescriptorAPointer = Marshal.AllocHGlobal(fileGroupDescriptorBytes.Length)
                        Marshal.Copy(fileGroupDescriptorBytes, 0, fileGroupDescriptorAPointer, fileGroupDescriptorBytes.Length)
                        Dim fileGroupDescriptorObject As Object = Marshal.PtrToStructure(fileGroupDescriptorAPointer, GetType(NativeMethods.FILEGROUPDESCRIPTORA))
                        Dim fileGroupDescriptor As NativeMethods.FILEGROUPDESCRIPTORA = CType(fileGroupDescriptorObject, NativeMethods.FILEGROUPDESCRIPTORA)
                        Dim fileNames As String() = New String(fileGroupDescriptor.cItems - 1) {}
                        Dim fileDescriptorPointer As IntPtr = CType((CInt(fileGroupDescriptorAPointer) + Marshal.SizeOf(fileGroupDescriptorAPointer)), IntPtr)

                        For fileDescriptorIndex As Integer = 0 To fileGroupDescriptor.cItems - 1
                            Dim fileDescriptor As NativeMethods.FILEDESCRIPTORA = CType(Marshal.PtrToStructure(fileDescriptorPointer, GetType(NativeMethods.FILEDESCRIPTORA)), NativeMethods.FILEDESCRIPTORA)
                            fileNames(fileDescriptorIndex) = fileDescriptor.cFileName
                            fileDescriptorPointer = CType((CInt(fileDescriptorPointer) + Marshal.SizeOf(fileDescriptor)), IntPtr)
                        Next

                        Return fileNames
                    Finally
                        Marshal.FreeHGlobal(fileGroupDescriptorAPointer)
                    End Try

                Case "FileGroupDescriptorW"
                    Dim fileGroupDescriptorWPointer As IntPtr = IntPtr.Zero

                    Try
                        Dim fileGroupDescriptorStream As MemoryStream = CType(Me.underlyingDataObject.GetData("FileGroupDescriptorW"), MemoryStream)
                        Dim fileGroupDescriptorBytes As Byte() = New Byte(fileGroupDescriptorStream.Length - 1) {}
                        fileGroupDescriptorStream.Read(fileGroupDescriptorBytes, 0, fileGroupDescriptorBytes.Length)
                        fileGroupDescriptorStream.Close()
                        fileGroupDescriptorWPointer = Marshal.AllocHGlobal(fileGroupDescriptorBytes.Length)
                        Marshal.Copy(fileGroupDescriptorBytes, 0, fileGroupDescriptorWPointer, fileGroupDescriptorBytes.Length)
                        Dim fileGroupDescriptorObject As Object = Marshal.PtrToStructure(fileGroupDescriptorWPointer, GetType(NativeMethods.FILEGROUPDESCRIPTORW))
                        Dim fileGroupDescriptor As NativeMethods.FILEGROUPDESCRIPTORW = CType(fileGroupDescriptorObject, NativeMethods.FILEGROUPDESCRIPTORW)
                        Dim fileNames As String() = New String(fileGroupDescriptor.cItems - 1) {}
                        Dim fileDescriptorPointer As IntPtr = CType((CInt(fileGroupDescriptorWPointer) + Marshal.SizeOf(fileGroupDescriptorWPointer)), IntPtr)

                        For fileDescriptorIndex As Integer = 0 To fileGroupDescriptor.cItems - 1
                            Dim fileDescriptor As NativeMethods.FILEDESCRIPTORW = CType(Marshal.PtrToStructure(fileDescriptorPointer, GetType(NativeMethods.FILEDESCRIPTORW)), NativeMethods.FILEDESCRIPTORW)
                            fileNames(fileDescriptorIndex) = fileDescriptor.cFileName
                            fileDescriptorPointer = CType((CInt(fileDescriptorPointer) + Marshal.SizeOf(fileDescriptor)), IntPtr)
                        Next

                        Return fileNames
                    Finally
                        Marshal.FreeHGlobal(fileGroupDescriptorWPointer)
                    End Try

                Case "FileContents"
                    Dim fileContentNames As String() = CType(Me.GetData("FileGroupDescriptor"), String())
                    Dim fileContents As MemoryStream() = New MemoryStream(fileContentNames.Length - 1) {}

                    For fileIndex As Integer = 0 To fileContentNames.Length - 1
                        fileContents(fileIndex) = Me.GetData(format, fileIndex)
                    Next

                    Return fileContents
            End Select

            Return Me.underlyingDataObject.GetData(format, autoConvert)
        End Function

        Public Function GetData(ByVal format As String, ByVal index As Integer) As MemoryStream ' Implements System.Windows.Forms.IDataObject.GetData
            Dim formatetc As FORMATETC = New FORMATETC()
            formatetc.cfFormat = CShort(DataFormats.GetFormat(format).Id)
            formatetc.dwAspect = DVASPECT.DVASPECT_CONTENT
            formatetc.lindex = index
            formatetc.ptd = New IntPtr(0)
            formatetc.tymed = TYMED.TYMED_ISTREAM Or TYMED.TYMED_ISTORAGE Or TYMED.TYMED_HGLOBAL
            Dim medium As STGMEDIUM = New STGMEDIUM()
            Me.comUnderlyingDataObject.GetData(formatetc, medium)

            Select Case medium.tymed
                Case TYMED.TYMED_ISTORAGE
                    Dim iStorage As NativeMethods.IStorage = Nothing
                    Dim iStorage2 As NativeMethods.IStorage = Nothing
                    Dim iLockBytes As NativeMethods.ILockBytes = Nothing
                    Dim iLockBytesStat As System.Runtime.InteropServices.ComTypes.STATSTG

                    Try
                        iStorage = CType(Marshal.GetObjectForIUnknown(medium.unionmember), NativeMethods.IStorage)
                        Marshal.Release(medium.unionmember)
                        iLockBytes = NativeMethods.CreateILockBytesOnHGlobal(IntPtr.Zero, True)
                        iStorage2 = NativeMethods.StgCreateDocfileOnILockBytes(iLockBytes, &H1012, 0)
                        iStorage.CopyTo(0, Nothing, IntPtr.Zero, iStorage2)
                        iLockBytes.Flush()
                        iStorage2.Commit(0)
                        iLockBytesStat = New System.Runtime.InteropServices.ComTypes.STATSTG()
                        iLockBytes.Stat(iLockBytesStat, 1)
                        Dim iLockBytesSize As Integer = CInt(iLockBytesStat.cbSize)
                        Dim iLockBytesContent As Byte() = New Byte(iLockBytesSize - 1) {}
                        iLockBytes.ReadAt(0, iLockBytesContent, iLockBytesContent.Length, Nothing)
                        Return New MemoryStream(iLockBytesContent)
                    Finally
                        Marshal.ReleaseComObject(iStorage2)
                        Marshal.ReleaseComObject(iLockBytes)
                        Marshal.ReleaseComObject(iStorage)
                    End Try

                Case TYMED.TYMED_ISTREAM
                    Dim iStream As IStream = Nothing
                    Dim iStreamStat As System.Runtime.InteropServices.ComTypes.STATSTG

                    Try
                        iStream = CType(Marshal.GetObjectForIUnknown(medium.unionmember), IStream)
                        Marshal.Release(medium.unionmember)
                        iStreamStat = New System.Runtime.InteropServices.ComTypes.STATSTG()
                        iStream.Stat(iStreamStat, 0)
                        Dim iStreamSize As Integer = CInt(iStreamStat.cbSize)
                        Dim iStreamContent As Byte() = New Byte(iStreamSize - 1) {}
                        iStream.Read(iStreamContent, iStreamContent.Length, IntPtr.Zero)
                        Return New MemoryStream(iStreamContent)
                    Finally
                        Marshal.ReleaseComObject(iStream)
                    End Try

                Case TYMED.TYMED_HGLOBAL
                    Return CType(Me.getDataFromHGLOBLALMethod.Invoke(Me.oleUnderlyingDataObject, New Object() {DataFormats.GetFormat(CShort(formatetc.cfFormat)).Name, medium.unionmember}), MemoryStream)
            End Select

            Return Nothing
        End Function

        Public Function GetDataPresent(ByVal format As Type) As Boolean Implements System.Windows.Forms.IDataObject.GetDataPresent
            Return Me.underlyingDataObject.GetDataPresent(format)
        End Function

        Public Function GetDataPresent(ByVal format As String) As Boolean Implements System.Windows.Forms.IDataObject.GetDataPresent
            Return Me.underlyingDataObject.GetDataPresent(format)
        End Function

        Public Function GetDataPresent(ByVal format As String, ByVal autoConvert As Boolean) As Boolean Implements System.Windows.Forms.IDataObject.GetDataPresent
            Return Me.underlyingDataObject.GetDataPresent(format, autoConvert)
        End Function

        Public Function GetFormats() As String() Implements System.Windows.Forms.IDataObject.GetFormats
            Return Me.underlyingDataObject.GetFormats()
        End Function

        Public Function GetFormats(ByVal autoConvert As Boolean) As String() Implements System.Windows.Forms.IDataObject.GetFormats
            Return Me.underlyingDataObject.GetFormats(autoConvert)
        End Function

        Public Sub SetData(ByVal data As Object) Implements System.Windows.Forms.IDataObject.SetData
            Me.underlyingDataObject.SetData(data)
        End Sub

        Public Sub SetData(ByVal format As Type, ByVal data As Object) Implements System.Windows.Forms.IDataObject.SetData
            Me.underlyingDataObject.SetData(format, data)
        End Sub

        Public Sub SetData(ByVal format As String, ByVal data As Object) Implements System.Windows.Forms.IDataObject.SetData
            Me.underlyingDataObject.SetData(format, data)
        End Sub

        Public Sub SetData(ByVal format As String, ByVal autoConvert As Boolean, ByVal data As Object) Implements System.Windows.Forms.IDataObject.SetData
            Me.underlyingDataObject.SetData(format, autoConvert, data)
        End Sub
    End Class
#End Region
End Namespace
