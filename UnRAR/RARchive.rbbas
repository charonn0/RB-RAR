#tag Class
Class RARchive
	#tag Method, Flags = &h0
		Function ArchiveComment() As String
		  Dim mHandle As Integer
		  Dim data As RAROpenArchiveData
		  Dim path As New MemoryBlock(ArchiveFile.AbsolutePath.LenB * 2)
		  path.CString(0) = ArchiveFile.AbsolutePath
		  Dim mb As New MemoryBlock(260 * 2)
		  data.AchiveName = path
		  data.OpenMode = UnRAR.RAR_OM_EXTRACT
		  data.CommentBufferSize = mb.Size
		  data.Comments = mb
		  mHandle = UnRAR.RAROpenArchive(data)
		  If mHandle <= 0 Then
		    Dim err As New RuntimeException
		    err.ErrorNumber = data.OpenResult
		    err.Message = UnRAR.FormatError(err.ErrorNumber)
		    Raise err
		  End If
		  Call UnRAR.RARCloseArchive(mHandle)
		  Dim comment As MemoryBlock = data.Comments
		  Return comment.StringValue(0, data.CommentSize)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Close()
		  If CallbackRefs.HasKey(Me.ArchiveFile.AbsolutePath) Then CallbackRefs.Remove(Me.ArchiveFile.AbsolutePath)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub Constructor(ArchFile As FolderItem)
		  mArchiveFile = ArchFile
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ExtractAll(OutputDirectory As FolderItem) As Boolean
		  Dim count As Integer = Me.ItemCount
		  Dim mhandle As Integer = Me.Handle(UnRAR.RAR_OM_EXTRACT)
		  If DecryptPassword <> "" Then
		    UnRAR.RARSetPassword(mHandle, DecryptPassword)
		  End If
		  Dim success As Boolean
		  Dim dir As MemoryBlock = OutputDirectory.AbsolutePath
		  For i As Integer = 0 To count
		    Dim h As RARHeaderData
		    If UnRAR.RARReadHeader(mHandle, h) <> 0 And UnRAR.RARProcessFile(mHandle, UnRAR.RAR_OM_EXTRACT, dir, Nil) <> 0 Then
		      success = False
		      Exit For
		    Else
		      success = True
		    End If
		  Next
		  Call UnRAR.RARCloseArchive(mHandle)
		  Return success
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ExtractItem(Index As Integer, OutputFile As FolderItem, Password As String = "") As Boolean
		  Dim count As Integer = ItemCount
		  Dim mhandle As Integer = Me.Handle(UnRAR.RAR_OM_EXTRACT)
		  If Password <> "" Then
		    UnRAR.RARSetPassword(mHandle, Password)
		  End If
		  Dim success As Boolean
		  Dim dir As MemoryBlock = OutputFile.AbsolutePath
		  For i As Integer = 0 To count
		    Dim h As RARHeaderData
		    If i = Index Then
		      If UnRAR.RARReadHeader(mHandle, h) <> 0 And UnRAR.RARProcessFile(mHandle, UnRAR.RAR_OM_EXTRACT, Nil, dir) <> 0 Then
		        success = False
		      Else
		        success = True
		      End If
		      Exit For
		    Else
		      If UnRAR.RARReadHeader(mHandle, h) <> 0 And UnRAR.RARProcessFile(mHandle, UnRAR.RAR_OM_LIST, Nil, Nil) <> 0 Then
		        success = False
		        Exit For
		      Else
		        success = True
		      End If
		    End If
		    
		  Next
		  Call UnRAR.RARCloseArchive(mHandle)
		  Return success
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Handle(mode As Integer) As Integer
		  Dim mHandle As Integer
		  Dim data As RAROpenArchiveData
		  Dim path As New MemoryBlock(ArchiveFile.AbsolutePath.LenB * 2)
		  path.CString(0) = ArchiveFile.AbsolutePath
		  data.AchiveName = path
		  data.OpenMode = mode
		  mHandle = UnRAR.RAROpenArchive(data)
		  If mHandle <= 0 Then
		    Dim err As New RuntimeException
		    err.ErrorNumber = data.OpenResult
		    err.Message = UnRAR.FormatError(err.ErrorNumber)
		    Raise err
		  End If
		  Return mHandle
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function HandleRAREvent(msg As Integer, UserData As Ptr, P1 As Ptr, P2 As Ptr) As Integer
		  Select Case msg
		  Case UnRAR.UCM_CHANGEVOLUME
		    Dim name As MemoryBlock = P1
		    Dim callmode As Integer = Integer(p2)
		    Select Case callmode
		    Case UnRAR.RAR_VOL_ASK
		      Dim nextvolume As FolderItem
		      nextvolume = RaiseEvent LoadNextVolume(name)
		      If nextvolume <> Nil And nextvolume.Exists Then
		        Return 0
		      Else
		        Return -1
		      End If
		    Case UnRAR.RAR_VOL_NOTIFY
		      If RaiseEvent NextVolumeLoaded(name) Then
		        Return -1
		      Else
		        Return 0
		      End If
		    End Select
		  Case UnRAR.UCM_NEEDPASSWORD
		    If DecryptPassword <> "" Then
		      Dim mb As MemoryBlock = DecryptPassword
		      p1.Ptr(0) = mb
		      p2.Int32(0) = mb.Size
		      Return 0
		    Else
		      mLastError = UnRAR.ErrorNeedPassword
		      Return -1
		    End If
		    
		  Case UnRAR.UCM_PROCESSDATA
		    
		  End Select
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Item(Index As Integer) As RARItem
		  Dim count As Integer = Me.ItemCount
		  Dim mhandle As Integer = Me.Handle(UnRAR.RAR_OM_EXTRACT)
		  Dim header As RARHeaderData
		  For i As Integer = 0 To count
		    mLastError = UnRAR.RARProcessFile(mHandle, UnRAR.RAR_OM_LIST, Nil, Nil)
		    If Me.LastError <> 0 Then Exit For
		    mLastError = UnRAR.RARReadHeader(mHandle, header)
		    If Me.LastError <> 0 Or Index = i Then
		      Exit For
		    End If
		  Next
		  Call UnRAR.RARCloseArchive(mHandle)
		  Return New RARItem(header, Index)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ItemCount() As Integer
		  Dim mhandle As Integer = Me.Handle(UnRAR.RAR_OM_LIST)
		  Dim count As Integer
		  While True
		    Dim header As RARHeaderData
		    mLastError = UnRAR.RARReadHeader(mHandle, header)
		    If LastError <> 0 Then Exit While
		    mLastError = UnRAR.RARProcessFile(mHandle, UnRAR.RAR_OM_LIST, Nil, Nil)
		    If LastError <> 0 Then Exit While
		    count = count + 1
		  Wend
		  Call UnRAR.RARCloseArchive(mHandle)
		  Return count
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function LastError() As Integer
		  Return mLastError
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		 Shared Function Open(Archive As FolderItem) As RARchive
		  Dim r As New RARchive(Archive)
		  CallbackRefs.Value(Archive.AbsolutePath) = r
		  Return r
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Shared Function RARCallback(msg As Integer, UserData As Ptr, P1 As Ptr, P2 As Ptr) As Integer
		  If CallbackRefs = Nil Then CallbackRefs = New Dictionary
		  Dim mb As MemoryBlock = UserData
		  Dim arc As FolderItem = GetFolderItem(mb.CString(0))
		  If CallbackRefs.HasKey(arc) Then
		    Dim rar As RARchive = CallbackRefs.Value(arc.AbsolutePath)
		    Return rar.HandleRAREvent(msg, UserData, P1, P2)
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Test() As Boolean
		  Dim count As Integer = Me.ItemCount
		  Dim mhandle As Integer = Me.Handle(UnRAR.OpTest)
		  Dim success As Boolean
		  For i As Integer = 0 To count
		    If UnRAR.RARProcessFile(mHandle, UnRAR.OpTest, Nil, Nil) <> 0 Then
		      success = False
		      Exit For
		    Else
		      success = True
		    End If
		  Next
		  Call UnRAR.RARCloseArchive(mHandle)
		  Return success
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function TestItem(Index As Integer) As Boolean
		  Dim count As Integer = Me.ItemCount
		  Dim mhandle As Integer = Me.Handle(UnRAR.OpTest)
		  Dim success As Boolean
		  If count < Index Then Raise New OutOfBoundsException
		  For i As Integer = 0 To Index - 1
		    If UnRAR.RARProcessFile(mHandle, UnRAR.OpTest, Nil, Nil) <> 0 Then
		      success = False
		      Exit For
		    ElseIf UnRAR.RARProcessFile(mHandle, UnRAR.RAR_OM_LIST, Nil, Nil) <> 0 Then
		      success = True
		    Else
		      success = False
		      Exit For
		    End If
		  Next
		  Call UnRAR.RARCloseArchive(mHandle)
		  Return success
		End Function
	#tag EndMethod


	#tag Hook, Flags = &h0
		Event LoadNextVolume(VolumeName As String) As FolderItem
	#tag EndHook

	#tag Hook, Flags = &h0
		Event NextVolumeLoaded(VolumeName As String) As Boolean
	#tag EndHook


	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  return mArchiveFile
			End Get
		#tag EndGetter
		ArchiveFile As FolderItem
	#tag EndComputedProperty

	#tag Property, Flags = &h21
		Private Shared CallbackRefs As Dictionary
	#tag EndProperty

	#tag Property, Flags = &h0
		DecryptPassword As String
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mArchiveFile As FolderItem
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mLastError As Integer
	#tag EndProperty


	#tag ViewBehavior
		#tag ViewProperty
			Name="DecryptPassword"
			Group="Behavior"
			Type="String"
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
