#tag Class
Class RARchive
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
		    CloseArchive(mHandle)
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
		  Return UBound(Me.ListItems) + 1
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ExtractItem(Index As Integer, SaveTo As FolderItem, Password As String = "") As Boolean
		  ' extracts the archived file at Index to SaveTo
		  ' Pass -1 and a directory to extract all items into the directory.
		  If UnRAR.IsAvailable Then
		    mLastError = 0
		    Dim mhandle As Integer = OpenArchive(RARFile, RAR_OM_EXTRACT)
		    If mhandle <= 0 Then mLastError = mhandle * -1
		    If Password <> "" Then RARSetPassword(mHandle, Password)
		    Dim i, mmode As Integer
		    Do Until Me.LastError <> 0
		      Dim header As RARHeaderData
		      mLastError = RARReadHeader(mHandle, header)
		      If mLastError <> 0 Then Continue
		      Dim pitem As New RARItem(header, i, RARFile)
		      Dim FilePath, DirPath As MemoryBlock
		      If i = Index Then
		        mmode = RAR_EXTRACT
		        RaiseEvent ProcessItem(pitem, mmode, SaveTo)
		        FilePath = New MemoryBlock(SaveTo.AbsolutePath.LenB * 2)
		        FilePath.CString(0) = SaveTo.AbsolutePath + Chr(0)
		      ElseIf Index = -1 Then
		        mmode = RAR_EXTRACT
		        If Not SaveTo.Directory Then SaveTo = SaveTo.Parent
		        RaiseEvent ProcessItem(pitem, mmode, SaveTo)
		        DirPath = New MemoryBlock(SaveTo.AbsolutePath.LenB * 2)
		        DirPath.CString(0) = SaveTo.AbsolutePath + Chr(0)
		      Else
		        mmode = RAR_SKIP
		        RaiseEvent ProcessItem(pitem, mmode, SaveTo)
		      End If
		      If FilePath = Nil Then FilePath = ""
		      If DirPath = Nil Then DirPath = ""
		      mLastError = RARProcessFile(mhandle, mmode, DirPath, FilePath)
		      
		      i = i + 1
		    Loop
		    CloseArchive(mHandle)
		    If Me.LastError = UnRAR.ErrorEndArchive Then mLastError = 0
		    Return mLastError = 0
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Item(Index As Integer) As RARItem
		  ' Retrieves the header for a single item
		  If UnRAR.IsAvailable Then
		    Dim mhandle As Integer = OpenArchive(RARFile, RAR_OM_LIST)
		    If mhandle < 0 Then mLastError = mhandle * -1
		    Dim header As RARHeaderData
		    Dim ritem As RARItem
		    Dim i As Integer
		    
		    Do Until Me.LastError <> 0
		      mLastError = RARReadHeader(mHandle, header)
		      If Index = i And mLastError = 0 Then
		        ritem = New RARItem(header, i, Me.RARFile)
		        Exit Do
		      ElseIf mLastError = 0 Then
		        mLastError = RARProcessFile(mHandle, RAR_SKIP, Nil, Nil)
		      End If
		      i = i + 1
		    Loop
		    CloseArchive(mHandle)
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

	#tag Method, Flags = &h0
		Function ListItems() As RARItem()
		  ' Returns the number of items in the archive
		  If UnRAR.IsAvailable Then
		    Dim items() As RARItem
		    mLastError = 0
		    Dim mhandle As Integer = OpenArchive(RARFile, RAR_OM_LIST)
		    If mhandle <= 0 Then mLastError = mhandle * -1
		    Dim count As Integer
		    Do Until Me.LastError <> 0
		      Dim header As RARHeaderData
		      mLastError = RARReadHeader(mHandle, header)
		      If mLastError = 0 Then
		        items.Append(New RARItem(header, count, RARFile))
		        mLastError = RARProcessFile(mHandle, RAR_SKIP, Nil, Nil)
		        count = count + 1
		      End If
		    Loop
		    CloseArchive(mHandle)
		    If Me.LastError = ErrorEndArchive Then mLastError = 0 ' not an error
		    Return items
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function RARFile() As FolderItem
		  Return mRARFile
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function TestItem(Index As Integer, Password As String = "") As Boolean
		  ' tests item(s) in the archive
		  ' Pass the index of the item to test, or pass -1 to test all items.
		  If UnRAR.IsAvailable Then
		    mLastError = 0
		    Dim mhandle As Integer = OpenArchive(RARFile, RAR_OM_EXTRACT)
		    If mhandle <= 0 Then mLastError = mhandle * -1
		    If Password <> "" Then RARSetPassword(mHandle, Password)
		    Dim i, mmode As Integer
		    Do Until Me.LastError <> 0
		      Dim header As RARHeaderData
		      mLastError = RARReadHeader(mHandle, header)
		      If mLastError <> 0 Then Continue
		      Dim pitem As New RARItem(header, i, RARFile)
		      Dim f As FolderItem
		      If i = Index Or Index = -1 Then
		        mmode = RAR_TEST
		      Else
		        mmode = RAR_SKIP
		      End If
		      RaiseEvent ProcessItem(pitem, mmode, f)
		      mLastError = RARProcessFile(mhandle, mmode, Nil, Nil)
		      i = i + 1
		    Loop
		    CloseArchive(mHandle)
		    If Me.LastError = UnRAR.ErrorEndArchive Then mLastError = 0
		    Return mLastError = 0
		  End If
		End Function
	#tag EndMethod


	#tag Hook, Flags = &h0
		Event ProcessItem(Item As RARItem, Operation As Integer, ByRef ExtractTo As FolderItem)
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
		instance can have the archive open (for extraction, header reading, testing, or counting) at any given moment.
		
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
