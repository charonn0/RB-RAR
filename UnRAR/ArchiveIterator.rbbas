#tag Class
Protected Class ArchiveIterator
	#tag Method, Flags = &h0
		Sub Close()
		  mCurrentItem = Nil
		  mCurrentIndex = -1
		  If mHandle <> 0 Then
		    CloseArchive(mHandle)
		  End If
		  mHandle = 0
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(ArchFile As FolderItem, Mode As Integer, Password As String)
		  mRARFile = ArchFile
		  OpenMode = Mode
		  mPassword = Password
		  Me.Reset
		  If mLastError <> 0 Then
		    Dim err As New IOException
		    err.ErrorNumber = mLastError
		    err.Message = UnRAR.FormatError(mLastError)
		    Raise err
		  End If
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(RARHandle As Integer, ArchFile As FolderItem, Mode As Integer, Password As String)
		  mHandle = RARHandle
		  Me.Constructor(ArchFile, Mode, Password)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function CurrentIndex() As Integer
		  Return mCurrentIndex
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function CurrentItem() As UnRAR.RARItem
		  Return mCurrentItem
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Destructor()
		  Me.Close
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Handle() As Integer
		  Return mHandle
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function LastError() As Integer
		  Return mLastError
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ProcessItem(ProcessingMode As Integer, ExtractPath As FolderItem = Nil) As Boolean
		  ' Calling this method will apply the ProcessingMode to the CurrentItem and advance the CurrentItem by one.
		  ' Returns true if ProcessingMode was applied and the next item was loaded. Returns False on error or end-of-archive;
		  ' check ArchiveIterator.LastError for details.
		  ' If ProcessingMode is RAR_EXTRACT then pass a FolderItem to extract into. Pass a directory to extract the item into
		  ' the directory using the item name (automatically creating subdirectories as needed); pass a non-existing file to extract
		  ' directly into the file.
		  
		  Dim FilePath, DirPath As MemoryBlock
		  
		  Select Case True
		  Case ExtractPath = Nil
		    FilePath = Nil
		    DirPath = Nil
		  Case ExtractPath.Directory And ExtractPath.Exists
		    DirPath = New MemoryBlock(ExtractPath.AbsolutePath.LenB * 2)
		    DirPath.CString(0) = ExtractPath.AbsolutePath + Chr(0)
		  Else
		    FilePath = New MemoryBlock(ExtractPath.AbsolutePath.LenB * 2)
		    FilePath.CString(0) = ExtractPath.AbsolutePath + Chr(0)
		  End Select
		  If FilePath = Nil Then FilePath = ""
		  If DirPath = Nil Then DirPath = ""
		  mLastError = RARProcessFile(mhandle, ProcessingMode, DirPath, FilePath)
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

	#tag Method, Flags = &h0
		Function RARFile() As FolderItem
		  Return mRARFile
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Reset()
		  Me.Close
		  mhandle = OpenArchive(RARFile, OpenMode)
		  ' >0 is a valid handle, <0 is a RAR error *-1
		  If mhandle <= 0 Then
		    mLastError = mhandle * -1
		    mHandle = 0
		    Return
		  ElseIf mPassword <> "" Then
		    RARSetPassword(mHandle, mPassword)
		  End If
		  Dim header As RARHeaderData
		  mLastError = RARReadHeader(mHandle, header)
		  If mLastError = 0 Then
		    mCurrentIndex = 0
		    mCurrentItem = New RARItem(header, mCurrentIndex, mRARFile)
		  End If
		  
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h1
		Protected mCurrentIndex As Integer
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mCurrentItem As UnRAR.RARItem
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mHandle As Integer
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mLastError As Integer
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mPassword As String
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mRARFile As FolderItem
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected OpenMode As Integer
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
