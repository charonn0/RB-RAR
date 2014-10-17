#tag Class
Protected Class ArchiveIterator
	#tag Method, Flags = &h0
		Sub Constructor(ArchFile As FolderItem, Mode As Integer = RAR_OM_EXTRACT, Password As String = "")
		  mRARFile = ArchFile
		  mLastError = 0
		  mhandle = OpenArchive(RARFile, RAR_OM_EXTRACT)
		  ' >0 is a valid handle, <0 is a RAR error *-1
		  If mhandle <= 0 Then
		    mLastError = mhandle * -1
		  ElseIf Password <> "" Then
		    RARSetPassword(mHandle, Password)
		  End If
		  Dim header As RARHeaderData
		  mLastError = RARReadHeader(mHandle, header)
		  If mLastError <> 0 Then
		    Dim err As New IOException
		    err.ErrorNumber = mLastError
		    err.Message = UnRAR.FormatError(mLastError)
		    Raise err
		  End If
		  mCurrentItem = New RARItem(header, mCurrentIndex, ArchFile)
		  
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
		  CloseArchive(mHandle)
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
