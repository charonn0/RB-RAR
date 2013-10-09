#tag Class
Class RARchive
	#tag Method, Flags = &h1
		Protected Shared Sub CloseArchive(RARHandle As Integer)
		  If UnRAR.IsAvailable Then
		    Call RARCloseArchive(RARHandle)
		  End If
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(RARFile As FolderItem)
		  mRARFile = RARFile
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ExtractAll(OutputDirectory As FolderItem, Password As String = "") As Boolean
		  ' extracts the entire archive to the specified output folder
		  If UnRAR.IsAvailable Then
		    mLastError = 0
		    Dim mhandle As Integer = OpenArchive(RARFile, RAR_OM_EXTRACT)
		    If mhandle <= 0 Then mLastError = mhandle * -1
		    Dim success As Boolean
		    If Password <> "" Then RARSetPassword(mHandle, Password)
		    Dim dir As MemoryBlock = OutputDirectory.AbsolutePath + Chr(0) + Chr(0)
		    
		    Do Until Me.LastError <> 0
		      Dim h As RARHeaderData
		      mLastError = UnRAR.RARProcessFile(mHandle, RAR_EXTRACT, dir, Nil)
		      success = (mLastError = 0)
		      If Not success Then Continue
		      mLastError = UnRAR.RARReadHeader(mHandle, h)
		      If h.FileName.Trim <> "" Then success = True
		    Loop
		    CloseArchive(mHandle)
		    Return success
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ExtractItem(Index As Integer, Password As String = "") As FolderItem
		  ' extracts the archived file at Index to a Temporary file, or Nil on error.
		  If UnRAR.IsAvailable Then
		    mLastError = 0
		    Dim mhandle As Integer = OpenArchive(RARFile, RAR_OM_EXTRACT)
		    If mhandle <= 0 Then mLastError = mhandle * -1
		    If Password <> "" Then RARSetPassword(mHandle, Password)
		    Dim savedto As FolderItem
		    Dim i As Integer
		    Do Until Me.LastError <> 0
		      Dim header As RARHeaderData
		      mLastError = RARReadHeader(mHandle, header)
		      If i = Index Then
		        If mLastError = 0 Then
		          Dim head As New RARItem(header, i)
		          savedto = SpecialFolder.Temporary.Child(NthField(head.FileName, "\", CountFields(head.FileName, "\")))
		          Dim path As New MemoryBlock(savedto.AbsolutePath.LenB * 2)
		          path.CString(0) = savedto.AbsolutePath + Chr(0)
		          mLastError = RARProcessFile(mHandle, RAR_EXTRACT, Nil, path)
		          If mLastError <> 0 Then savedto = Nil
		          Exit Do
		        End If
		      Else
		        mLastError = RARProcessFile(mHandle, RAR_SKIP, Nil, Nil)
		      End If
		      i = i + 1
		    Loop
		    CloseArchive(mHandle)
		    Return savedto
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Item(Index As Integer) As RARItem
		  ' Retrieves the header for a single item
		  Dim mhandle As Integer = OpenArchive(RARFile, RAR_OM_EXTRACT)
		  Dim header As RARHeaderData
		  Dim ritem As RARItem
		  Dim i As Integer
		  If UnRAR.IsAvailable Then
		    Do Until Me.LastError <> 0
		      mLastError = RARReadHeader(mHandle, header)
		      If Index = i Then
		        ritem = New RARItem(header, i)
		        Exit Do
		      Else
		        mLastError = RARProcessFile(mHandle, RAR_SKIP, Nil, Nil)
		      End If
		      i = i + 1
		    Loop
		    CloseArchive(mHandle)
		    Return ritem
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function LastError() As Integer
		  Return mLastError
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Shared Function OpenArchive(RARFile As FolderItem, Mode As Integer, ExtendedMode As Boolean = False) As Integer
		  If UnRAR.IsAvailable Then
		    Dim mHandle, err As Integer
		    Dim mb As New MemoryBlock(260 * 2)
		    Dim path As New MemoryBlock(RARFile.AbsolutePath.LenB * 2)
		    path.CString(0) = RARFile.AbsolutePath
		    If Not ExtendedMode Then
		      Dim data As RAROpenArchiveData
		      data.CommentBufferSize = mb.Size
		      data.Comments = mb
		      data.AchiveName = path
		      data.OpenMode = mode
		      mHandle = UnRAR.RAROpenArchive(data)
		      err = data.OpenResult
		    Else
		      Dim data As RAROpenArchiveDataEx
		      data.CommentBufferSize = mb.Size
		      data.Comments = mb
		      data.AchiveName = path
		      data.OpenMode = mode
		      data.Callback = AddressOf RARCallbackHandler
		      data.UserData = path
		      mHandle = RAROpenArchiveEx(data)
		      err = data.OpenResult
		    End If
		    If mHandle <= 0 Then Return err * -1
		    Return mHandle
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Shared Function RARCallbackHandler(msg As UInt32, UserData As Ptr, P1 As Ptr, P2 As Ptr) As Integer
		  Dim s As String = UserData.CString(0)
		  If ExArchives.HasKey(s) Then
		    Dim a As RARchive = ExArchives.Value(s)
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
		    If RaiseEvent ProcessData(bs, bs.Length) Then
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
		Function RARFile() As FolderItem
		  Return mRARFile
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function TestAll(Password As String = "") As Boolean
		  ' tests all items in the archive
		  ' tests a single item in the archive
		  If UnRAR.IsAvailable Then
		    Dim i, count As Integer
		    count = Me.Count
		    Dim mhandle As Integer = OpenArchive(RARFile, RAR_OM_EXTRACT)
		    If Password <> "" Then RARSetPassword(mHandle, Password)
		    Dim success As Boolean
		    Do Until Me.LastError <> 0
		      If i > count Then
		        Exit Do
		      Else
		        mLastError = RARProcessFile(mHandle, RAR_TEST, Nil, Nil)
		      End If
		      success = (Me.LastError = 0)
		      i = i + 1
		    Loop
		    CloseArchive(mHandle)
		    Return success
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function TestItem(Index As Integer, Password As String = "") As Boolean
		  ' tests a single item in the archive
		  If UnRAR.IsAvailable Then
		    Dim mhandle As Integer = OpenArchive(RARFile, RAR_OM_EXTRACT)
		    If Password <> "" Then RARSetPassword(mHandle, Password)
		    Dim success As Boolean
		    Dim i As Integer
		    Do Until Me.LastError <> 0
		      If i = Index Then
		        mLastError = RARProcessFile(mHandle, RAR_TEST, Nil, Nil)
		        Exit Do
		      Else
		        mLastError = RARProcessFile(mHandle, RAR_SKIP, Nil, Nil)
		      End If
		      success = (Me.LastError = 0)
		      i = i + 1
		    Loop
		    CloseArchive(mHandle)
		    Return success
		  End If
		End Function
	#tag EndMethod


	#tag Hook, Flags = &h0
		Event ChangeVolume(ByRef NextVolume As FolderItem) As Boolean
	#tag EndHook

	#tag Hook, Flags = &h0
		Event PasswordPrompt() As String
	#tag EndHook

	#tag Hook, Flags = &h0
		Event ProcessData(Data As Readable, Length As Integer) As Boolean
	#tag EndHook


	#tag Note, Name = About this class
		This class represents a RAR archive. Pass the RAR as a FolderItem to the class constructor. To access
		files inside the archive call the ExtractAll or ExtractItem methods. 
		
		To retreive metadata for a particular item, call the Item method. The Item method returns a RARItem 
		for the archived file at the specified Index.
		
		To test a single file in the archive, call TestItem; to test all files, call TestAll.
		
		Indices passed to Item, ExtractItem, and TestItem are zero-based: they run from 0 to RARchive.Count-1.
		
		The Comment method returns the archive comment, if any.
		
		You can have multiple instances of the RARchive class pointing to the same RAR file. However, only one
		instance can have the archive open (for extraction, testing, or counting) at any given moment.
		
		Avoid unneccessary calls to RARchive.Count as each call must enumerate the contents of the entire archive. Similarly,
		the execution time of RARchive.Item rises in direct proportion to the Index of the file being retrieved. Patterns
		and optimizations appropriate for enumerating a directory via the FolderItem class are equally appropriate for 
		RARchives.
	#tag EndNote


	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
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
			End Get
		#tag EndGetter
		Comment As String
	#tag EndComputedProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  ' Returns the number of items in the archive
			  If UnRAR.IsAvailable Then
			    mLastError = 0
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
			    If Me.LastError = UnRAR.ErrorEndArchive Then mLastError = 0
			    Return count
			  End If
			End Get
		#tag EndGetter
		Count As Integer
	#tag EndComputedProperty

	#tag ComputedProperty, Flags = &h1
		#tag Getter
			Get
			  Static mExArchives As Dictionary
			  If mExArchives = Nil Then mExArchives = New Dictionary
			  Return mExArchives
			End Get
		#tag EndGetter
		Protected Shared ExArchives As Dictionary
	#tag EndComputedProperty

	#tag Property, Flags = &h0
		ExtendedMode As Boolean
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mLastError As Integer
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mRARFile As FolderItem
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
