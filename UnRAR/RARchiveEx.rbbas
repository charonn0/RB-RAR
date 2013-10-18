#tag Class
Class RARchiveEx
Inherits RARchive
	#tag Method, Flags = &h0
		Sub Close()
		  If ExArchives.HasKey(Me) Then
		    ExArchives.Remove(Me)
		  End If
		  CloseArchive(mHandle)
		  mLastError = 0
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Shared Function ExArchives() As Dictionary
		  Static mExArchives As Dictionary
		  If mExArchives = Nil Then mExArchives = New Dictionary
		  Return mExArchives
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ExtractAll()
		  ' extracts the entire archive to the specified output folder
		  If UnRAR.IsAvailable Then
		    mhandle = OpenArchiveEx(RARFile, RAR_OM_EXTRACT)
		    If mhandle <= 0 Then mLastError = mhandle * -1
		    ExArchives.Value(RARFile.AbsolutePath) = Me
		    mIndex = 0
		    Do Until LastError <> 0
		      mLastError = UnRAR.RARProcessFile(mHandle, RAR_EXTRACT, Nil, Nil)
		      If mLastError = 0 Then mIndex = mIndex + 1
		    Loop
		    Me.Close
		  End If
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ExtractItem(Index As Integer)
		  ' extracts the archived file at Index to a Temporary file, or Nil on error.
		  If UnRAR.IsAvailable Then
		    mhandle = OpenArchiveEx(RARFile, RAR_OM_EXTRACT)
		    If mhandle <= 0 Then mLastError = mhandle * -1
		    ExArchives.Value(RARFile.AbsolutePath) = Me
		    mIndex = 0
		    Do Until Me.LastError <> 0
		      Dim header As RARHeaderDataEx
		      mLastError = RARReadHeaderEx(mHandle, header)
		      If mIndex = Index Then
		        If mLastError = 0 Then
		          mLastError = RARProcessFileW(mHandle, RAR_EXTRACT, Nil, Nil)
		          Exit Do
		        End If
		      Else
		        mLastError = RARProcessFileW(mHandle, RAR_SKIP, Nil, Nil)
		      End If
		      mIndex = mIndex + 1
		    Loop
		    Me.Close
		  End If
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Flags() As Integer
		  Dim info As RAROpenArchiveDataEx = GetArchiveInfoEx(RARFile)
		  Return info.Flags
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Shared Function GetArchiveInfoEx(RARFile As FolderItem) As RAROpenArchiveDataEx
		  If UnRAR.IsAvailable Then
		    Dim mHandle, err As Integer
		    Dim mb As New MemoryBlock(260 * 2)
		    Dim path As New MemoryBlock(RARFile.AbsolutePath.LenB * 2)
		    path.CString(0) = RARFile.AbsolutePath
		    Dim data As RAROpenArchiveDataEx
		    data.CommentBufferSize = mb.Size
		    data.Comments = mb
		    data.AchiveName = path
		    data.OpenMode = RAR_OM_EXTRACT
		    data.Callback = AddressOf RARCallbackHandler
		    data.UserData = path
		    mHandle = RAROpenArchiveEx(data)
		    If mHandle <= 0 Then data.OpenResult = err * -1
		    CloseArchive(mHandle)
		    Return data
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function HasAuthenticityInfo() As Boolean
		  Return BitAnd(Me.Flags, ArchiveFlag_AuthenticityInfo) = ArchiveFlag_AuthenticityInfo
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function HasRecoveryRecord() As Boolean
		  Return BitAnd(Me.Flags, ArchiveFlag_RecoveryRecord) = ArchiveFlag_RecoveryRecord
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function IsEncrypted() As Boolean
		  Return BitAnd(Me.Flags, ArchiveFlag_EncryptedNames) = ArchiveFlag_EncryptedNames
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function IsLocked() As Boolean
		  Return BitAnd(Me.Flags, ArchiveFlag_Locked) = ArchiveFlag_Locked
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function IsSolid() As Boolean
		  Return BitAnd(Me.Flags, ArchiveFlag_Solid) = ArchiveFlag_Solid
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Item(Index As Integer)
		  ' Retrieves the header for a single item
		  If UnRAR.IsAvailable Then
		    Dim mhandle As Integer = OpenArchiveEx(RARFile, RAR_OM_EXTRACT)
		    ExArchives.Value(RARFile.AbsolutePath) = Me
		    Dim header As RARHeaderDataEx
		    Dim ritem As RARItemEx
		    Dim i As Integer
		    
		    Do Until Me.LastError <> 0
		      mLastError = RARReadHeaderEx(mHandle, header)
		      If Index = i Then
		        ritem = New RARItemEx(header, i)
		        Exit Do
		      Else
		        mLastError = RARProcessFileW(mHandle, RAR_SKIP, Nil, Nil)
		      End If
		      i = i + 1
		    Loop
		    Me.Close
		  End If
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Shared Function OpenArchiveEx(RARFile As FolderItem, Mode As Integer) As Integer
		  If UnRAR.IsAvailable Then
		    Dim mHandle, err As Integer
		    Dim mb As New MemoryBlock(260 * 2)
		    Dim path As New MemoryBlock(RARFile.AbsolutePath.LenB * 2)
		    path.CString(0) = RARFile.AbsolutePath
		    Dim data As RAROpenArchiveDataEx
		    data.CommentBufferSize = mb.Size
		    data.Comments = mb
		    data.AchiveName = path
		    data.OpenMode = mode
		    data.Callback = AddressOf RARCallbackHandler
		    data.UserData = path
		    mHandle = RAROpenArchiveEx(data)
		    err = data.OpenResult
		    If mHandle <= 0 Then Return err * -1
		    Return mHandle
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Shared Function RARCallbackHandler(msg As UInt32, UserData As Ptr, P1 As Ptr, P2 As Ptr) As Integer
		  #pragma X86CallingConvention StdCall
		  Dim mb As MemoryBlock = UserData
		  Dim s As String = mb.CString(0)
		  If ExArchives.HasKey(s) Then
		    Dim a As RARchiveEx = ExArchives.Value(s)
		    Return a.RAREventHandler(msg, P1, P2)
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function RAREventHandler(msg As UInt32, P1 As Ptr, P2 As Ptr) As Integer
		  Select Case msg
		  Case UCM_CHANGEVOLUME
		    If P2.UInt32(0) = RAR_VOL_ASK Then
		      Dim path As String = P1.CString(0)
		      Dim f As FolderItem = GetFolderItem(path)
		      If RaiseEvent ChangeVolume(f) Then
		        P1.CString(0) = f.AbsolutePath + Chr(0)
		        Return 1
		      End If
		    ElseIf P2.UInt32(0) = RAR_VOL_NOTIFY Then
		      Return 1
		    End If
		    
		  Case UCM_PROCESSDATA
		    Dim mb As MemoryBlock = P1.CString(0)
		    Dim bs As New BinaryStream(mb)
		    Dim h As RARHeaderDataEx
		    mLastError = UnRAR.RARReadHeaderEx(mHandle, h)
		    If RaiseEvent ProcessData(bs, bs.Length, New RARItemEx(h, mIndex)) Then
		      Return 1
		    End If
		    
		  Case UCM_NEEDPASSWORD
		    Dim pw As String = RaiseEvent PasswordPrompt()
		    If pw.Trim <> "" Then
		      pw = ConvertEncoding(pw, Encodings.ASCII)
		      P1.CString(0) = pw + Chr(0)
		      P2.UInt32(0) = pw.Len
		      Return 1
		    End If
		    
		  End Select
		  
		  Return -1
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub TestAll()
		  ' tests all items in the archive
		  If UnRAR.IsAvailable Then
		    Dim count As Integer
		    count = Me.Count
		    Dim mhandle As Integer = OpenArchiveEx(RARFile, RAR_OM_EXTRACT)
		    ExArchives.Value(RARFile.AbsolutePath) = Me
		    mIndex = 0
		    Do Until Me.LastError <> 0
		      If mIndex > count Then
		        Exit Do
		      Else
		        mLastError = RARProcessFileW(mHandle, RAR_TEST, Nil, Nil)
		      End If
		      mIndex = mIndex + 1
		    Loop
		    CloseArchive(mHandle)
		  End If
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub TestItem(Index As Integer)
		  ' tests a single item in the archive
		  If UnRAR.IsAvailable Then
		    Dim mhandle As Integer = OpenArchiveEx(RARFile, RAR_OM_EXTRACT)
		    ExArchives.Value(RARFile.AbsolutePath) = Me
		    mIndex = 0
		    Do Until Me.LastError <> 0
		      If mIndex = Index Then
		        mLastError = RARProcessFileW(mHandle, RAR_TEST, Nil, Nil)
		        Exit Do
		      Else
		        mLastError = RARProcessFileW(mHandle, RAR_SKIP, Nil, Nil)
		      End If
		      mIndex = mIndex + 1
		    Loop
		    Me.Close
		  End If
		End Sub
	#tag EndMethod


	#tag Hook, Flags = &h0
		Event ChangeVolume(ByRef NextVolume As FolderItem) As Boolean
	#tag EndHook

	#tag Hook, Flags = &h0
		Event PasswordPrompt() As String
	#tag EndHook

	#tag Hook, Flags = &h0
		Event ProcessData(Data As Readable, Length As Integer, Header As RARItemEx) As Boolean
	#tag EndHook


	#tag Property, Flags = &h21
		Private mHandle As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mIndex As Integer
	#tag EndProperty


	#tag ViewBehavior
		#tag ViewProperty
			Name="Comment"
			Group="Behavior"
			Type="String"
			InheritedFrom="RARchive"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Count"
			Group="Behavior"
			Type="Integer"
			InheritedFrom="RARchive"
		#tag EndViewProperty
		#tag ViewProperty
			Name="ExtendedMode"
			Group="Behavior"
			Type="Boolean"
			InheritedFrom="RARchive"
		#tag EndViewProperty
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
