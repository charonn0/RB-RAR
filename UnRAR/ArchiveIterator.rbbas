#tag Class
Protected Class ArchiveIterator
	#tag Method, Flags = &h0
		Function ArchiveComment() As String
		  Return mCommentBuffer.Trim
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Close()
		  mCurrentItem = Nil
		  mCurrentIndex = 0
		  If mHandle <> 0 Then mLastError = RARCloseArchive(mHandle)
		  mHandle = 0
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub Constructor(RARHandle As Integer, ArchFile As FolderItem, Mode As Integer, Password As String)
		  mHandle = RARHandle
		  mRARFile = ArchFile
		  mOpenMode = Mode
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

	#tag Method, Flags = &h1
		Protected Sub LoadCurrentHeader()
		  Dim header As RARHeaderData
		  mLastError = RARReadHeader(mHandle, header)
		  If mLastError = 0 Then
		    mCurrentItem = New RARItem(header, mCurrentIndex, RARFile)
		    mCurrentIndex = mCurrentIndex + 1
		  End If
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		 Shared Function Open(RARFile As FolderItem, Mode As Integer, Password As String) As UnRAR.ArchiveIterator
		  If UnRAR.IsAvailable Then
		    Dim h As Integer
		    Dim mb As New MemoryBlock(260 * 2)
		    h = ReOpen(RARFile, Mode, mb)
		    If h > 0 Then
		      Dim a As New ArchiveIterator(h, RARFile, Mode, Password)
		      a.mCommentBuffer = mb
		      Return a
		    End If
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function OpenMode() As Integer
		  Return mOpenMode
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
		      mCurrentIndex = mCurrentIndex + 1
		      mCurrentItem = New RARItem(header, mCurrentIndex, RARFile)
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

	#tag Method, Flags = &h1
		Protected Shared Function ReOpen(RARFile As FolderItem, Mode As Integer, CommentBuffer As MemoryBlock) As Integer
		  If UnRAR.IsAvailable Then
		    Dim h, err As Integer
		    Dim path As New MemoryBlock(RARFile.AbsolutePath.LenB * 2)
		    path.CString(0) = RARFile.AbsolutePath
		    Dim data As RAROpenArchiveData
		    data.CommentBufferSize = CommentBuffer.Size
		    data.Comments = CommentBuffer
		    data.ArchiveName = path
		    data.OpenMode = mode
		    h = UnRAR.RAROpenArchive(data)
		    err = data.OpenResult
		    If h > 0 Then
		      Return h
		    Else
		      Return err * -1
		    End If
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Reset()
		  Me.Close
		  mCommentBuffer = New MemoryBlock(260 * 2)
		  mhandle = ReOpen(mRARFile, mOpenMode, mCommentBuffer)
		  ' >0 is a valid handle, <0 is a RAR error *-1
		  If mhandle <= 0 Then
		    mLastError = mhandle * -1
		    mHandle = 0
		    Return
		  ElseIf mPassword <> "" Then
		    RARSetPassword(mHandle, mPassword)
		  End If
		  LoadCurrentHeader()
		  
		End Sub
	#tag EndMethod


	#tag Note, Name = About this class
		This class more closely resembles the UnRAR library interface than the RARchive class.
		
		Create a new instance, and access the CurrentItem to read the headers of the currently selected item in the archive.
		
		Call the ProcessItem method to process the current item and then advance the selection to the next item in the archive.
	#tag EndNote


	#tag Property, Flags = &h1
		Protected mCommentBuffer As MemoryBlock
	#tag EndProperty

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
		Protected mOpenMode As Integer
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mPassword As String
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
