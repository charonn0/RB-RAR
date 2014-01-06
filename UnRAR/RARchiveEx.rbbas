#tag Class
Class RARchiveEx
	#tag Method, Flags = &h0
		Sub Close()
		  If UnRAR.IsAvailable And RARHandle > 0 Then
		    Call RARCloseArchive(RARHandle)
		  End If
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(RARFile As FolderItem, Password As String = "")
		  mRARFile = RARFile
		  
		  If UnRAR.IsAvailableEx Then
		    ArchivePassword = Password.Trim
		    Dim path As New MemoryBlock(2048)
		    path.CString(0) = RARFile.AbsolutePath
		    Dim mb As New MemoryBlock(260 * 2)
		    ArchiveHeader.ArchiveName = path
		    ArchiveHeader.OpenMode = UnRAR.RAR_OM_EXTRACT
		    ArchiveHeader.CommentBufferSize = mb.Size
		    ArchiveHeader.Comments = mb
		    ArchiveHeader.Callback = AddressOf RARCallbackHandler
		    Dim data As New MemoryBlock(4)
		    data.Int32Value(0) = Path.CRC32
		    ArchiveHeader.UserData = data
		    ExArchives.Value(Path.CRC32) = New WeakRef(Me)
		    RARHandle = UnRAR.RAROpenArchiveEx(ArchiveHeader)
		    If RARHandle <= 0 Then
		      mLastError = ArchiveHeader.OpenResult
		      CloseArchive(RARHandle)
		      Dim err As New IOException
		      err.ErrorNumber = mLastError
		      err.Message = "Unable to open RAR file."
		      Raise err
		    End If
		  Else
		    mLastError = ErrorRARUnavailable
		  End If
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
		Function ExtractAll(SaveTo As FolderItem) As Boolean
		  If UnRAR.IsAvailable Then
		    mLastError = 0
		    Dim FilePath As New MemoryBlock(SaveTo.AbsolutePath.LenB * 2 + 2)
		    FilePath.WString(0) = SaveTo.AbsolutePath' + Chr(0)
		    Do Until Me.LastError <> 0
		      Dim header As RARHeaderDataEx
		      mLastError = RARReadHeaderEx(RARHandle, header)
		      If mLastError = 0 Then
		        If SaveTo.Directory Then
		          mLastError = RARProcessFileW(RARHandle, RAR_EXTRACT, FilePath, Nil)
		        Else
		          mLastError = RARProcessFileW(RARHandle, RAR_EXTRACT, Nil, FilePath)
		        End If
		      End If
		    Loop
		    If Me.LastError = ErrorEndArchive Then mLastError = 0 ' not an error
		    Return mLastError = 0
		  Else
		    mLastError = ErrorRARUnavailable
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function LastError() As Integer
		  Return mLastError
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ListItems() As RARItem()
		  ' Returns a list of all items in the archive
		  If UnRAR.IsAvailable Then
		    Dim items() As RARItem
		    Dim count As Integer
		    mLastError = 0
		    Do Until Me.LastError <> 0
		      Dim header As RARHeaderDataEx
		      mLastError = RARReadHeaderEx(RARHandle, header)
		      If mLastError = 0 Then
		        items.Append(New RARItem(header, count, RARFile))
		        mLastError = RARProcessFileW(RARHandle, RAR_SKIP, Nil, Nil)
		        count = count + 1
		      End If
		    Loop
		    If Me.LastError = ErrorEndArchive Then mLastError = 0 ' not an error
		    Return items
		  Else
		    mLastError = ErrorRARUnavailable
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Shared Function RARCallbackHandler(msg As UInt32, UserData As Ptr, P1 As Ptr, P2 As Ptr) As Integer
		  #pragma X86CallingConvention StdCall
		  Dim mb As MemoryBlock = UserData
		  Dim s As Integer = mb.Int32Value(0)
		  If ExArchives.HasKey(s) Then
		    Dim w As WeakRef = ExArchives.Value(s)
		    If w.Value <> Nil Then
		      Dim r As RARchiveEx = RARchiveEx(w.Value)
		      Return r.RAREventHandler(msg, P1, P2)
		    End If
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function RAREventHandler(msg As UInt32, ByRef P1 As Ptr, ByRef P2 As Ptr) As Integer
		  Select Case msg
		  Case UCM_CHANGEVOLUME, UCM_CHANGEVOLUMEW
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
		    Return 1
		    
		  Case UCM_NEEDPASSWORD, UCM_NEEDPASSWORDW
		    'Dim pw As String = ConvertEncoding(ArchivePassword, Encodings.ASCII) + Chr(0)
		    P1.CString(0) = ArchivePassword + Chr(0)
		    'P2.Int32(0) = pw.Len
		    Return 1
		  End Select
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function RARFile() As FolderItem
		  Return mRARFile
		End Function
	#tag EndMethod


	#tag Hook, Flags = &h0
		Event ChangeVolume(ByRef NextVolume As FolderItem) As Boolean
	#tag EndHook


	#tag Property, Flags = &h1
		Protected ArchiveHeader As RAROpenArchiveDataEx
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected ArchivePassword As String
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mIndex As Integer = 0
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mLastError As Integer
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mRARFile As FolderItem
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected RARHandle As Integer
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
