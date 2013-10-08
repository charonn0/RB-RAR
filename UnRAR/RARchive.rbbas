#tag Class
Protected Class RARchive
	#tag Method, Flags = &h1
		Protected Shared Function ArchiveHeader(RARFile As FolderItem) As RAROpenArchiveData
		  ' Retrieves the header for the entire archive
		  If UnRAR.IsAvailable Then
		    Dim mHandle As Integer
		    Dim data As RAROpenArchiveData
		    Dim path As New MemoryBlock(RARFile.AbsolutePath.LenB * 2)
		    path.CString(0) = RARFile.AbsolutePath
		    data.AchiveName = path
		    data.OpenMode = RAR_OM_EXTRACT
		    mHandle = UnRAR.RAROpenArchive(data)
		    If mHandle > 0 Then
		      CloseArchive(mHandle)
		      Return data
		    End If
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Shared Sub CloseArchive(RARHandle As Integer)
		  If UnRAR.IsAvailable Then
		    Call RARCloseArchive(RARHandle)
		  End If
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Comment() As String
		  Dim mHandle As Integer
		  Dim data As RAROpenArchiveData
		  Dim path As New MemoryBlock(RARFile.AbsolutePath.LenB * 2)
		  path.CString(0) = RARFile.AbsolutePath
		  Dim mb As New MemoryBlock(260 * 2)
		  data.AchiveName = path
		  data.OpenMode = UnRAR.RAR_OM_EXTRACT
		  data.CommentBufferSize = mb.Size
		  data.Comments = mb
		  mHandle = UnRAR.RAROpenArchive(data)
		  If mHandle <= 0 Then
		    mLastError = data.OpenResult
		    CloseArchive(mHandle)
		    Return ""
		  End If
		  CloseArchive(mHandle)
		  Dim comment As MemoryBlock = data.Comments
		  Return comment.StringValue(0, data.CommentSize)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(RARFile As FolderItem)
		  If UnRAR.IsRARArchive(RARFile) Then
		    mRARFile = RARFile
		  Else
		    mLastError = UnRAR.ErrorBadArchive
		  End If
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Count() As Integer
		  ' Returns the number of items in the archive
		  If UnRAR.IsAvailable Then
		    Dim mhandle As Integer = OpenArchive(RARFile, RAR_OM_LIST)
		    Dim count As Integer
		    While True
		      Dim header As RARHeaderData
		      mLastError = UnRAR.RARProcessFile(mHandle, RAR_SKIP, Nil, Nil)
		      If mLastError <> 0 Then Exit While
		      mLastError = UnRAR.RARReadHeader(mHandle, header)
		      If mLastError <> 0 Then Exit While
		      count = count + 1
		    Wend
		    CloseArchive(mHandle)
		    Return count
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ExtractAll(OutputDirectory As FolderItem) As Boolean
		  ' extracts the entire archive to the specified output folder
		  If UnRAR.IsAvailable Then
		    Dim mhandle As Integer = OpenArchive(RARFile, RAR_OM_EXTRACT)
		    Dim success As Boolean
		    If Password <> "" Then
		      UnRAR.RARSetPassword(mHandle, Password)
		    End If
		    Dim dir As MemoryBlock = OutputDirectory.AbsolutePath + Chr(0) + Chr(0)
		    While True
		      Dim h As RARHeaderData
		      mLastError = UnRAR.RARProcessFile(mHandle, RAR_EXTRACT, dir, Nil)
		      success = (mLastError = 0)
		      If Not success Then Exit While
		      mLastError = UnRAR.RARReadHeader(mHandle, h)
		      If h.FileName.Trim <> "" Then
		        success = True
		      Else
		        Exit While
		      End If
		    Wend
		    CloseArchive(mHandle)
		    Return success
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ExtractItem(Index As Integer) As FolderItem
		  ' extracts the archived file at Index to a Temporary file
		  If UnRAR.IsAvailable Then
		    Dim mhandle As Integer = OpenArchive(RARFile, RAR_OM_EXTRACT)
		    If Password <> "" Then
		      UnRAR.RARSetPassword(mHandle, Password)
		    End If
		    Dim outputfile As FolderItem
		    Dim i As Integer
		    While True
		      Dim h As RARHeaderData
		      If i < Index Then
		        If UnRAR.RARProcessFile(mHandle, RAR_SKIP, Nil, Nil) <> 0 Then
		          outputfile = Nil
		          Exit While
		        End If
		      ElseIf i = Index Then
		        Call UnRAR.RARReadHeader(mHandle, h)
		        If h.FileName.Trim <> "" Then
		          outputfile = SpecialFolder.Temporary.Child(h.FileName)
		          Dim bs As BinaryStream = BinaryStream.Create(outputfile, True)
		          bs.Close
		        End If
		        Dim file As MemoryBlock = outputfile.AbsolutePath + Chr(0) + Chr(0)
		        If UnRAR.RARProcessFile(mHandle, RAR_EXTRACT, Nil, file) <> 0 Then
		          outputfile = Nil
		          Exit While
		        Else
		          Exit While
		        End If
		      Else
		        Exit While
		      End If
		      i = i + 1
		    Wend
		    CloseArchive(mHandle)
		    Return outputfile
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Item(Index As Integer) As RARItem
		  ' Retrieves the header for a single item
		  Dim mhandle As Integer = OpenArchive(RARFile, RAR_OM_EXTRACT)
		  Dim header As RARHeaderData
		  Dim i As Integer
		  If UnRAR.IsAvailable Then
		    While True
		      If UnRAR.RARProcessFile(mHandle, RAR_SKIP, Nil, Nil) <> 0 Then Exit While
		      If UnRAR.RARReadHeader(mHandle, header)<> 0 Or Index = i Then Exit While
		      i = i + 1
		    Wend
		    CloseArchive(mHandle)
		    If i = Index Then
		      Return New RARItem(header, i)
		    End If
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function LastError() As Integer
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Shared Function OpenArchive(RARFile As FolderItem, Mode As Integer) As Integer
		  If UnRAR.IsAvailable Then
		    Dim mHandle As Integer
		    Dim data As RAROpenArchiveData
		    Dim path As New MemoryBlock(RARFile.AbsolutePath.LenB * 2)
		    path.CString(0) = RARFile.AbsolutePath
		    data.AchiveName = path
		    data.OpenMode = mode
		    mHandle = UnRAR.RAROpenArchive(data)
		    If mHandle <= 0 Then Return 0
		    Return mHandle
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function TestAll() As Boolean
		  ' tests all items in the archive
		  If UnRAR.IsAvailable Then
		    Dim i, count As Integer
		    count = Me.Count
		    Dim mhandle As Integer = OpenArchive(RARFile, RAR_OM_EXTRACT)
		    Dim success As Boolean
		    While True
		      If i <= Count Then
		        If RARProcessFile(mHandle, RAR_TEST, Nil, Nil) <> 0 Then
		          success = False
		          Exit While
		        Else
		          success = True
		        End If
		      Else
		        Exit While
		      End If
		      i = i + 1
		    Wend
		    CloseArchive(mHandle)
		    Return success
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function TestItem(Index As Integer) As Boolean
		  ' tests a single item in the archive
		  If UnRAR.IsAvailable Then
		    Dim mhandle As Integer = OpenArchive(RARFile, RAR_OM_EXTRACT)
		    Dim success As Boolean
		    Dim i As Integer
		    While True
		      If i = Index Then
		        If RARProcessFile(mHandle, RAR_TEST, Nil, Nil) <> 0 Then
		          success = False
		        Else
		          success = True
		        End If
		        Exit While
		      ElseIf i < Index Then
		        If RARProcessFile(mHandle, RAR_SKIP, Nil, Nil) <> 0 Then
		          success = False
		          Exit While
		        Else
		          success = True
		        End If
		      Else
		        Exit While
		      End If
		      i = i + 1
		    Wend
		    CloseArchive(mHandle)
		    Return success
		  End If
		End Function
	#tag EndMethod


	#tag Property, Flags = &h1
		Protected mLastError As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mRARFile As FolderItem
	#tag EndProperty

	#tag Property, Flags = &h0
		Password As String
	#tag EndProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  return mRARFile
			End Get
		#tag EndGetter
		#tag Setter
			Set
			  mRARFile = value
			End Set
		#tag EndSetter
		RARFile As FolderItem
	#tag EndComputedProperty


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
			Name="Password"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
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
