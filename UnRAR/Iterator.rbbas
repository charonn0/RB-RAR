#tag Class
Protected Class Iterator
	#tag Method, Flags = &h0
		Sub Constructor(RARFile As FolderItem, Mode As Integer)
		  If Not UnRAR.IsAvailable Then Raise New PlatformNotSupportedException
		  mArchFile = RARFile
		  mOpenMode = Mode
		  mCommentBuffer = New MemoryBlock(260 * 2)
		  Dim path As New MemoryBlock(RARFile.AbsolutePath.LenB + 1)
		  path.CString(0) = RARFile.AbsolutePath
		  mArchiveHeader.CommentBufferSize = mCommentBuffer.Size
		  mArchiveHeader.Comments = mCommentBuffer
		  mArchiveHeader.ArchiveName = path
		  mArchiveHeader.OpenMode = mode
		  mHandle = RAROpenArchive(mArchiveHeader)
		  mLastError = mArchiveHeader.OpenResult
		  If mHandle = 0 Then
		    Raise New RuntimeException
		  End If
		  
		  Dim h As RARHeaderData
		  mLastError = RARReadHeader(mHandle, h)
		  If mLastError <> 0 Then Raise New RuntimeException
		  mCurrentIndex = 0
		  mCurrentItem = New UnRAR.ArchiveEntry(h, mCurrentIndex, mArchFile)
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function CurrentItem() As UnRAR.ArchiveEntry
		  Return mCurrentItem
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub Destructor()
		  If mHandle <> 0 Then mLastError = RARCloseArchive(mHandle)
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
		Function MoveNext(ProcessingMode As Integer, ExtractPath As FolderItem = Nil) As Boolean
		  ' Calling this method will apply the ProcessingMode to the CurrentItem and advance the CurrentItem by one.
		  ' Returns true if ProcessingMode was applied and the next item was loaded. Returns False on error or end-of-archive;
		  ' check ArchiveIterator.LastError for details.
		  ' If ProcessingMode is RAR_EXTRACT then pass a FolderItem to extract into. Pass a directory to extract the item into
		  ' the directory using the item name (automatically creating subdirectories as needed); pass a non-existing file to extract
		  ' directly into the file.
		  
		  Dim FilePath, DirPath As MemoryBlock
		  mLastError = 0
		  
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
		      mCurrentIndex = mCurrentIndex + 1
		      mCurrentItem = New UnRAR.ArchiveEntry(header, mCurrentIndex, mArchFile)
		    End If
		  End If
		  
		  Return mLastError = 0
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Reset()
		  Me.Destructor
		  Me.Constructor(mArchFile, mOpenMode)
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h0
		ArchivePassword As String
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mArchFile As FolderItem
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mArchiveHeader As RAROpenArchiveData
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mCommentBuffer As MemoryBlock
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mCurrentIndex As Integer
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mCurrentItem As UnRAR.ArchiveEntry
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mHandle As Integer
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mLastError As Integer
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mOpenMode As Integer
	#tag EndProperty


	#tag ViewBehavior
		#tag ViewProperty
			Name="ArchivePassword"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
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
