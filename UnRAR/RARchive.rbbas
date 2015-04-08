#tag Class
Class RARchive
	#tag Method, Flags = &h0
		Function Comment() As String
		  ' Returns the archive comment, if any
		  Dim rar As UnRAR.ArchiveIterator
		  rar = rar.Open(RARFile, UnRAR.RAR_OM_EXTRACT, "")
		  Return rar.ArchiveComment
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(RARFile As FolderItem)
		  mRARFile = RARFile
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Count() As Integer
		  '  Returns the number of items in the archive
		  Return UBound(Me.ListItems) + 1
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ExtractItem(Index As Integer, SaveTo As FolderItem, Password As String = "") As Boolean
		  ' extracts the archived file at Index to SaveTo
		  ' Pass -1 and a directory to extract all items into the directory.
		  
		  Dim rar As UnRAR.ArchiveIterator = UnRAR.IterateArchive(mRARFile, RAR_OM_EXTRACT, Password)
		  
		  Do Until rar.LastError <> 0
		    If RaiseEvent OperationProgress(rar.CurrentItem) Then
		      mLastError = UnRAR.ErrorUserCancel
		      Exit Do
		    End If
		    
		    Dim mmode As Integer
		    If rar.CurrentIndex = Index Then
		      mmode = RAR_EXTRACT
		    ElseIf Index = -1 Then
		      mmode = RAR_EXTRACT
		    Else
		      mmode = RAR_SKIP
		    End If
		    
		    If Not rar.ProcessItem(mmode, SaveTo) Then
		      mLastError = rar.LastError
		    ElseIf rar.CurrentIndex - 1 = Index Then
		      Exit Do
		    End If
		  Loop
		  
		  If Me.LastError = UnRAR.ErrorEndArchive Then
		    If rar.CurrentIndex < Index Or Index < -1 Then
		      Dim err As New OutOfBoundsException
		      err.Message = "RAR archive does not contain an entry at that index."
		      Raise err
		    End If
		    mLastError = 0
		  End If
		  
		  Return mLastError = 0
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Item(Index As Integer) As RARItem
		  ' Retrieves the header for a single item
		  Dim rar As UnRAR.ArchiveIterator = UnRAR.IterateArchive(mRARFile, RAR_OM_LIST)
		  Do Until rar.LastError <> 0
		    If RaiseEvent OperationProgress(rar.CurrentItem) Then
		      mLastError = UnRAR.ErrorUserCancel
		      Exit Do
		    End If
		    
		    If rar.CurrentIndex = Index Then Return rar.CurrentItem
		    If Not rar.ProcessItem(RAR_SKIP) Then
		      mLastError = rar.LastError
		    End If
		  Loop
		  
		  If Me.LastError = UnRAR.ErrorEndArchive Then
		    If rar.CurrentIndex < Index Or Index < -1 Then
		      Dim err As New OutOfBoundsException
		      err.Message = "RAR archive does not contain an entry at that index."
		      err.ErrorNumber = Index
		      Raise err
		    End If
		    mLastError = 0
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
		  Dim items() As RARItem
		  mLastError = 0
		  
		  Dim rar As UnRAR.ArchiveIterator = UnRAR.IterateArchive(mRARFile, RAR_OM_LIST, "")
		  Do Until rar.LastError <> 0
		    If RaiseEvent OperationProgress(rar.CurrentItem) Then
		      mLastError = UnRAR.ErrorUserCancel
		      Exit Do
		    End If
		    items.Append(rar.CurrentItem)
		    If Not rar.ProcessItem(RAR_SKIP) Then
		      mLastError = rar.LastError
		    End If
		  Loop
		  
		  If Me.LastError = ErrorEndArchive Then mLastError = 0 ' not an error
		  Return items
		  
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
		  Dim rar As UnRAR.ArchiveIterator = UnRAR.IterateArchive(mRARFile, RAR_OM_EXTRACT, Password)
		  
		  Do Until rar.LastError <> 0
		    If RaiseEvent OperationProgress(rar.CurrentItem) Then
		      mLastError = UnRAR.ErrorUserCancel
		      Exit Do
		    End If
		    
		    Dim mmode As Integer
		    If rar.CurrentIndex = Index Then
		      mmode = RAR_TEST
		    ElseIf Index = -1 Then
		      mmode = RAR_TEST
		    Else
		      mmode = RAR_SKIP
		    End If
		    
		    If Not rar.ProcessItem(mmode) Then
		      mLastError = rar.LastError
		    ElseIf rar.CurrentIndex - 1 = Index Then
		      Exit Do
		    End If
		  Loop
		  
		  If Me.LastError = UnRAR.ErrorEndArchive Then
		    If rar.CurrentIndex < Index Or Index < -1 Then
		      Dim err As New OutOfBoundsException
		      err.Message = "RAR archive does not contain an entry at that index."
		      Raise err
		    End If
		    mLastError = 0
		  End If
		  
		  Return mLastError = 0
		End Function
	#tag EndMethod


	#tag Hook, Flags = &h0
		Event OperationProgress(NextItem As RARItem) As Boolean
	#tag EndHook


	#tag Note, Name = About this class
		This class represents a RAR archive. Pass the RAR as a FolderItem to the class constructor. To access
		files inside the archive call the ExtractItem method. 
		
		This class can access archives with encrypted file data but not archives with encrypted file names.
		
		To retrieve metadata for a particular item, call the Item method. The Item method returns a RARItem 
		for the archived file at the specified Index.
		
		To test file(s), call TestItem.
		
		Indices passed to RARchive.Item, RARchive.ExtractItem, and RARchive.TestItem are zero-based: they run from 0 
		to RARchive.Count-1. Pass -1 as the Index to operate on all items in the archive. Use RARchive.ListItems() instead 
		of RARchive.Item(-1) or RARchive.Item() in a loop. 
		
		RARchive.TestItem and RARchive.ExtractItem raise the OperationProgress event once before each file is processed.
		Return True from the OperationProgress event to cancel any further processing.
		
		Performance optimizations appropriate for FolderItem.Item are equally effective with RARchive.Item, 
		RARchive.ExtractItem, and RARchive.TestItem
		
		The Comment method returns the archive comment, if any.
		
		You can have multiple instances of the RARchive class pointing to the same RAR file. However, only one
		instance can have the archive open (for extraction, header reading, testing, or counting) at any given moment.
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
