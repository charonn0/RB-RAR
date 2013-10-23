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
		Function Comment() As String
		  If UnRAR.IsAvailable Then
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
		    Dim comment As MemoryBlock = data.Comments
		    Dim sz As Integer = data.CommentSize
		    If mHandle > 0 Then CloseArchive(mHandle)
		    Return comment.StringValue(0, sz).Trim
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(RARFile As FolderItem)
		  mRARFile = RARFile
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Count() As Integer
		  ' Returns the number of items in the archive
		  If UnRAR.IsAvailable Then
		    mLastError = 0
		    Dim mhandle As Integer = OpenArchive(RARFile, RAR_OM_LIST)
		    If mhandle <= 0 Then mLastError = mhandle * -1
		    Dim count As Integer
		    Do Until Me.LastError <> 0
		      Dim header As RARHeaderData
		      mLastError = RARReadHeader(mHandle, header)
		      If mLastError = 0 Then
		        mLastError = RARProcessFile(mHandle, RAR_SKIP, Nil, Nil)
		        count = count + 1
		      End If
		    Loop
		    If mHandle > 0 Then CloseArchive(mHandle)
		    If Me.LastError = ErrorEndArchive Then mLastError = 0 ' not an error
		    Return count
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ExtractAll(OutputDirectory As FolderItem, Password As String = "") As Boolean
		  ' extracts the entire archive to the specified output folder
		  ' Use this method rather than iterating over the archive with
		  ' the ExtractItem method.
		  
		  If UnRAR.IsAvailable Then
		    mLastError = 0
		    Dim mhandle As Integer = OpenArchive(RARFile, RAR_OM_EXTRACT)
		    If mhandle <= 0 Then mLastError = mhandle * -1
		    If Password <> "" Then RARSetPassword(mHandle, Password)
		    Dim path As New MemoryBlock(OutputDirectory.AbsolutePath.LenB * 2)
		    path.CString(0) = OutputDirectory.AbsolutePath + Chr(0)
		    Do Until Me.LastError <> 0
		      Dim header As RARHeaderData
		      mLastError = RARReadHeader(mHandle, header)
		      If mLastError = 0 Then
		        'Dim head As New RARItem(header, i)
		        mLastError = RARProcessFile(mHandle, RAR_EXTRACT, path, Nil)
		      Else
		        If mLastError = UnRAR.ErrorEndArchive Then mLastError = 0
		        Exit Do
		      End If
		    Loop
		    If mHandle > 0 Then CloseArchive(mHandle)
		    Return mLastError = 0
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ExtractItem(Index As Integer, SaveTo As FolderItem, Password As String = "") As Boolean
		  ' extracts the archived file at Index
		  If UnRAR.IsAvailable Then
		    mLastError = 0
		    Dim mhandle As Integer = OpenArchive(RARFile, RAR_OM_EXTRACT)
		    If mhandle <= 0 Then mLastError = mhandle * -1
		    If Password <> "" Then RARSetPassword(mHandle, Password)
		    Dim i As Integer
		    Do Until Me.LastError <> 0
		      Dim header As RARHeaderData
		      mLastError = RARReadHeader(mHandle, header)
		      If i = Index Then
		        If mLastError = 0 Then
		          Dim path As New MemoryBlock(SaveTo.AbsolutePath.LenB * 2)
		          path.CString(0) = SaveTo.AbsolutePath + Chr(0)
		          mLastError = RARProcessFile(mHandle, RAR_EXTRACT, Nil, path)
		          Exit Do
		        End If
		      Else
		        mLastError = RARProcessFile(mHandle, RAR_SKIP, Nil, Nil)
		      End If
		      i = i + 1
		    Loop
		    If mHandle > 0 Then CloseArchive(mHandle)
		    If Me.LastError = UnRAR.ErrorEndArchive Then mLastError = 0
		    Return mLastError = 0
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Shared Function GetArchiveInfo(RARFile As FolderItem) As RAROpenArchiveData
		  If UnRAR.IsAvailable Then
		    Dim mHandle, err As Integer
		    Dim mb As New MemoryBlock(260 * 2)
		    Dim path As New MemoryBlock(RARFile.AbsolutePath.LenB * 2)
		    path.CString(0) = RARFile.AbsolutePath
		    Dim data As RAROpenArchiveData
		    data.CommentBufferSize = mb.Size
		    data.Comments = mb
		    data.AchiveName = path
		    data.OpenMode = RAR_OM_EXTRACT
		    mHandle = UnRAR.RAROpenArchive(data)
		    err = data.OpenResult
		    If mHandle <= 0 Then data.OpenResult = err * -1
		    CloseArchive(mHandle)
		    Return data
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Item(Index As Integer) As RARItem
		  ' Retrieves the header for a single item
		  If UnRAR.IsAvailable Then
		    Dim mhandle As Integer = OpenArchive(RARFile, RAR_OM_EXTRACT)
		    If mhandle < 0 Then mLastError = mhandle * -1
		    Dim header As RARHeaderData
		    Dim ritem As RARItem
		    Dim i As Integer
		    
		    Do Until Me.LastError <> 0
		      mLastError = RARReadHeader(mHandle, header)
		      If Index = i Then
		        ritem = New RARItem(header, i, Me.RARFile)
		        Exit Do
		      Else
		        mLastError = RARProcessFile(mHandle, RAR_SKIP, Nil, Nil)
		      End If
		      i = i + 1
		    Loop
		    If mHandle > 0 Then CloseArchive(mHandle)
		    If Me.LastError = UnRAR.ErrorEndArchive Then mLastError = 0
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
		Protected Shared Function OpenArchive(RARFile As FolderItem, Mode As Integer) As Integer
		  If UnRAR.IsAvailable Then
		    Dim mHandle, err As Integer
		    Dim mb As New MemoryBlock(260 * 2)
		    Dim path As New MemoryBlock(RARFile.AbsolutePath.LenB * 2)
		    path.CString(0) = RARFile.AbsolutePath
		    Dim data As RAROpenArchiveData
		    data.CommentBufferSize = mb.Size
		    data.Comments = mb
		    data.AchiveName = path
		    data.OpenMode = mode
		    mHandle = UnRAR.RAROpenArchive(data)
		    err = data.OpenResult
		    If mHandle <= 0 Then Return err * -1
		    Return mHandle
		  End If
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
		    mLastError = 0
		    Dim mhandle As Integer = OpenArchive(RARFile, RAR_OM_EXTRACT)
		    If mhandle <= 0 Then mLastError = mhandle * -1
		    If Password <> "" Then RARSetPassword(mHandle, Password)
		    
		    Do Until Me.LastError <> 0
		      Dim header As RARHeaderData
		      mLastError = RARReadHeader(mHandle, header)
		      If mLastError = 0 Then
		        mLastError = RARProcessFile(mHandle, RAR_TEST, Nil, Nil)
		      Else
		        If mLastError = UnRAR.ErrorEndArchive Then mLastError = 0
		        Exit Do
		      End If
		    Loop
		    If mHandle > 0 Then CloseArchive(mHandle)
		    If Me.LastError = UnRAR.ErrorEndArchive Then mLastError = 0
		    Return mLastError = 0
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function TestItem(Index As Integer, Password As String = "") As Boolean
		  ' tests a single item in the archive
		  If UnRAR.IsAvailable Then
		    mLastError = 0
		    Dim mhandle As Integer = OpenArchive(RARFile, RAR_OM_EXTRACT)
		    If mhandle <= 0 Then mLastError = mhandle * -1
		    If Password <> "" Then RARSetPassword(mHandle, Password)
		    Dim i As Integer
		    Do Until Me.LastError <> 0
		      Dim header As RARHeaderData
		      mLastError = RARReadHeader(mHandle, header)
		      If i = Index Then
		        If mLastError = 0 Then
		          mLastError = RARProcessFile(mHandle, RAR_TEST, Nil, Nil)
		          Exit Do
		        End If
		      Else
		        mLastError = RARProcessFile(mHandle, RAR_SKIP, Nil, Nil)
		      End If
		      i = i + 1
		    Loop
		    If mHandle > 0 Then CloseArchive(mHandle)
		    If Me.LastError = UnRAR.ErrorEndArchive Then mLastError = 0
		    Return mLastError = 0
		  End If
		End Function
	#tag EndMethod


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
