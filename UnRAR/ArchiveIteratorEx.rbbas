#tag Class
Protected Class ArchiveIteratorEx
Inherits UnRAR.ArchiveIterator
	#tag Method, Flags = &h1000
		Sub Constructor(ArchFile As FolderItem, Mode As Integer, Password As String)
		  If Instances = Nil Then Instances = New Dictionary
		  mRARFile = ArchFile
		  OpenMode = Mode
		  mPassword = Password
		  Me.Reset
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ProcessItem(ProcessingMode As Integer, ExtractPath As FolderItem = Nil) As Boolean
		  Dim FilePath, DirPath As MemoryBlock
		  
		  Select Case True
		  Case ExtractPath = Nil
		    FilePath = Nil
		    DirPath = Nil
		  Case ExtractPath.Directory And ExtractPath.Exists
		    DirPath = New MemoryBlock(ExtractPath.AbsolutePath.LenB * 2)
		    DirPath.WString(0) = ExtractPath.AbsolutePath + Chr(0)
		  Else
		    FilePath = New MemoryBlock(ExtractPath.AbsolutePath.LenB * 2)
		    FilePath.WString(0) = ExtractPath.AbsolutePath + Chr(0)
		  End Select
		  If FilePath = Nil Then FilePath = ""
		  If DirPath = Nil Then DirPath = ""
		  mLastError = RARProcessFileW(mhandle, ProcessingMode, DirPath, FilePath)
		  If mLastError = 0 Then
		    Dim header As RARHeaderData
		    mLastError = RARReadHeader(mHandle, header)
		    If mLastError = 0 Then
		      mCurrentItem = New RARItem(header, mCurrentIndex, RARFile)
		      mCurrentIndex = mCurrentIndex + 1
		    End If
		  End If
		  
		  Return mLastError = 0
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Shared Function RARCallbackHandler(msg As UInt32, UserData As Ptr, P1 As Ptr, P2 As Ptr) As Integer
		  '#pragma X86CallingConvention StdCall
		  Dim s As Integer = Integer(UserData)
		  Dim w As WeakRef = Instances.Lookup(s, Nil)
		  If w.Value <> Nil Then
		    Dim r As UnRAR.ArchiveIteratorEx = UnRAR.ArchiveIteratorEx(w.Value)
		    Return r.RAREventHandler(msg, P1, P2)
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function RAREventHandler(msg As UInt32, P1 As Ptr, P2 As Ptr) As Integer
		  Select Case msg
		  Case UCM_CHANGEVOLUME, UCM_CHANGEVOLUMEW
		    If P2.UInt32(0) = RAR_VOL_ASK Then
		      Dim f As FolderItem = GetFolderItem(P1.CString(0))
		      If RaiseEvent ChangeVolume(f) Then
		        P1.CString(0) = f.AbsolutePath + Chr(0)
		        Return 1
		      End If
		    ElseIf P2.UInt32(0) = RAR_VOL_NOTIFY Then
		      Dim f As FolderItem = GetFolderItem(P1.CString(0))
		      If Not RaiseEvent VolumeChanged(f) Then
		        Return 1
		      Else
		        Return -1 ' cancel
		      End If
		    End If
		    
		  Case UCM_PROCESSDATA
		    Dim sz As Integer = Integer(P2)
		    Dim mb As New MemoryBlock(sz)
		    Dim p As MemoryBlock = P1.Ptr(0)
		    mb.StringValue(0, mb.Size) = p.StringValue(0, mb.Size)
		    If Not RaiseEvent ProcessData(mb) Then
		      Return 1
		    Else
		      Return -1 'cancel
		    End If
		    
		  Case UCM_NEEDPASSWORD, UCM_NEEDPASSWORDW
		    If mPassword.Trim = "" Then mPassword = RaiseEvent PasswordRequired()
		    P1.CString(0) = mPassword + Chr(0)
		    Return 1
		  End Select
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Reset()
		  Me.Close
		  Dim path As New MemoryBlock(2048)
		  path.CString(0) = RARFile.AbsolutePath
		  Dim mb As New MemoryBlock(260 * 2)
		  ArchiveHeader.ArchiveName = path
		  ArchiveHeader.OpenMode = OpenMode
		  ArchiveHeader.CommentBufferSize = mb.Size
		  ArchiveHeader.Comments = mb
		  ArchiveHeader.Callback = AddressOf RARCallbackHandler
		  Dim ud As Integer = Microseconds
		  ArchiveHeader.UserData = Ptr(ud)
		  Instances.Value(ud) = New WeakRef(Me)
		  mHandle = UnRAR.RAROpenArchiveEx(ArchiveHeader)
		  If mHandle <= 0 Then
		    mLastError = ArchiveHeader.OpenResult
		    CloseArchive(mHandle)
		    Dim err As New IOException
		    err.ErrorNumber = mLastError
		    err.Message = "Unable to open RAR file."
		    Raise err
		  End If
		  Dim header As RARHeaderData
		  mLastError = RARReadHeader(mHandle, header)
		  If mLastError = 0 Then
		    mCurrentIndex = 0
		    mCurrentItem = New RARItem(header, mCurrentIndex, mRARFile)
		  End If
		  
		End Sub
	#tag EndMethod


	#tag Hook, Flags = &h0
		Event ChangeVolume(ByRef NextVolume As FolderItem) As Boolean
	#tag EndHook

	#tag Hook, Flags = &h0
		Event PasswordRequired() As String
	#tag EndHook

	#tag Hook, Flags = &h0
		Event ProcessData(NewData As MemoryBlock) As Boolean
	#tag EndHook

	#tag Hook, Flags = &h0
		Event VolumeChanged(CurrentVolume As FolderItem) As Boolean
	#tag EndHook


	#tag Property, Flags = &h1
		Protected ArchiveHeader As RAROpenArchiveDataEx
	#tag EndProperty

	#tag Property, Flags = &h21
		Private Shared Instances As Dictionary
	#tag EndProperty


	#tag ViewBehavior
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			InitialValue="-2147483648"
			InheritedFrom="Object"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=true
			Group="Position"
			InitialValue="0"
			InheritedFrom="Object"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			InheritedFrom="Object"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=true
			Group="ID"
			InheritedFrom="Object"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			InheritedFrom="Object"
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
