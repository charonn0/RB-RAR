#tag Class
Protected Class IteratorEx
Inherits UnRAR.Iterator
	#tag Method, Flags = &h0
		Sub Close()
		  Super.Close
		  mVolumeNumber = 1
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(RARFile As FolderItem, Mode As Integer = UnRAR.RAR_OM_EXTRACT, Password As String = "")
		  Super.Constructor(RARFile, Mode, Password)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function CurrentItem() As UnRAR.ArchiveEntry
		  Try
		    If Not mIsOpen Then OpenArchive()
		  Catch Err As RARException
		    Return Nil
		  End Try
		  Return Super.CurrentItem
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
		  ' If processing mode is RAR_TEST or RAR_SKIP then pass Nil as ExtractPath. Even in testing mode decompressed data will
		  ' be passed to the ProcessData event, allowing for decompression without writing to the disk. In skip mode no data is
		  ' passed to the ProcessData event.
		  
		  Dim FilePath, DirPath As MemoryBlock
		  mLastError = 0
		  Try
		    If Not mIsOpen Then OpenArchive()
		  Catch Err As RARException
		    Return False
		  End Try
		  
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
		    Dim header As RARHeaderDataEx
		    mLastError = RARReadHeaderEx(mHandle, header)
		    mCurrentIndex = mCurrentIndex + 1
		    If mLastError = 0 Then mCurrentItem = New UnRAR.ArchiveEntry(header, mCurrentIndex, Me.ArchiveFile)
		  End If
		  
		  Return mLastError = 0
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub OpenArchive()
		  ' Opens the archive and moves the current index to 0. Only one
		  ' instance may have an archive open at any given time.
		  
		  If Not UnRAR.IsAvailable Then Raise New PlatformNotSupportedException
		  mCommentBuffer = New MemoryBlock(260 * 2)
		  Dim path As New MemoryBlock(mArchFile.AbsolutePath.LenB * 2 + 2)
		  Static UserData As Integer
		  UserData = UserData + 1
		  Dim UserDatum As Integer = UserData
		  If Instances = Nil Then Instances = New Dictionary
		  Instances.Value(UserDatum) = New WeakRef(Me)
		  path.WString(0) = mArchFile.AbsolutePath
		  mArchiveHeader.CommentBufferSize = mCommentBuffer.Size
		  mArchiveHeader.Comments = mCommentBuffer
		  mArchiveHeader.ArchiveNameW = path
		  mArchiveHeader.Callback = AddressOf RARCallback
		  mArchiveHeader.UserData = UserDatum
		  mArchiveHeader.OpenMode = mOpenMode
		  
		  mHandle = RAROpenArchiveEx(mArchiveHeader)
		  mLastError = mArchiveHeader.OpenResult
		  If mHandle = 0 Then Raise New RARException(mLastError)
		  
		  
		  Dim h As RARHeaderDataEx
		  mLastError = RARReadHeaderEx(mHandle, h)
		  If mLastError <> 0 Then Raise New RARException(mLastError)
		  mCurrentIndex = 0
		  mCurrentItem = New UnRAR.ArchiveEntry(h, mCurrentIndex, mArchFile)
		  mIsOpen = True
		  mVolumeNumber = 1
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Shared Function RARCallback(Message As UInt32, UserData As Integer, Param1 As Ptr, Param2 As Ptr) As Integer
		  #pragma X86CallingConvention StdCall
		  If Instances = Nil Then Return -1
		  Dim RAR As WeakRef = Instances.Lookup(UserData, Nil)
		  If RAR <> Nil And RAR.Value <> Nil And RAR.Value IsA UnRAR.IteratorEx Then
		    Return UnRAR.IteratorEx(RAR.Value)._RARCallback(Message, Param1, Param2)
		  End If
		  Return -1
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Reset(StartIndex As Integer = 0)
		  Me.Close
		  For i As Integer = 0 To StartIndex - 1
		    If Not Me.MoveNext(RAR_SKIP) Then Exit For
		  Next
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function _RARCallback(Message As UInt32, Param1 As Ptr, Param2 As Ptr) As Integer
		  Select Case Message
		  Case UCM_CHANGEVOLUME, UCM_CHANGEVOLUMEW
		    Dim mb As MemoryBlock = Param1
		    Dim path As String
		    If Message = UCM_CHANGEVOLUMEW Then path = mb.WString(0) Else path = mb.CString(0)
		    Dim f As FolderItem = GetFolderItem(path)
		    
		    Select Case UInt32(Param2)
		    Case RAR_VOL_ASK
		      If RaiseEvent ChangeVolume(mVolumeNumber + 1, f) Then
		        If f <> Nil Then ' if f is nil (or unchanged) and the event returned true then UnRAR should retry the expected path
		          If Message = UCM_CHANGEVOLUMEW Then mb.WString(0) = f.AbsolutePath Else mb.CString(0) = f.AbsolutePath
		        End If
		        Return 1
		      End If
		      Return -1
		      
		    Case RAR_VOL_NOTIFY
		      If Message = UCM_CHANGEVOLUMEW Then 
		        mVolumeNumber = mVolumeNumber + 1
		        If Not RaiseEvent VolumeChanged(mVolumeNumber, f) Then Return 1
		        Return -1 ' abort
		      Else
		        Return 1
		      End If
		    End Select
		    
		  Case UCM_PROCESSDATA
		    If Not RaiseEvent ProcessData(Param1, Integer(Param2)) Then Return 1
		    Return -1 ' abort
		    
		  Case UCM_NEEDPASSWORD, UCM_NEEDPASSWORDW
		    If ArchivePassword.Trim = "" And Not RaiseEvent GetPassword(ArchivePassword) Then Return 0
		    Dim mb As MemoryBlock = Param1
		    If Message = UCM_NEEDPASSWORDW Then
		      mb.WString(0) = ArchivePassword
		    Else
		      mb.CString(0) = ArchivePassword
		    End If
		    Param2 = Ptr(ArchivePassword.Len)
		    Return 1
		    
		  End Select
		End Function
	#tag EndMethod


	#tag Hook, Flags = &h0
		Event ChangeVolume(VolumeNumber As Integer, ByRef NextVolume As FolderItem) As Boolean
	#tag EndHook

	#tag Hook, Flags = &h0
		Event GetPassword(ByRef ArchivePassword As String) As Boolean
	#tag EndHook

	#tag Hook, Flags = &h0
		Event ProcessData(NewData As MemoryBlock, Length As Integer) As Boolean
	#tag EndHook

	#tag Hook, Flags = &h0
		Event VolumeChanged(VolumeNumber As Integer, NewVolume As FolderItem) As Boolean
	#tag EndHook


	#tag Property, Flags = &h21
		Private Shared Instances As Dictionary
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mArchiveHeader As RAROpenArchiveDataEx
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mVolumeNumber As Integer = 1
	#tag EndProperty


	#tag Constant, Name = RAR_VOL_ASK, Type = Double, Dynamic = False, Default = \"0", Scope = Private
	#tag EndConstant

	#tag Constant, Name = RAR_VOL_NOTIFY, Type = Double, Dynamic = False, Default = \"1", Scope = Private
	#tag EndConstant

	#tag Constant, Name = UCM_CHANGEVOLUME, Type = Double, Dynamic = False, Default = \"0", Scope = Private
	#tag EndConstant

	#tag Constant, Name = UCM_CHANGEVOLUMEW, Type = Double, Dynamic = False, Default = \"3", Scope = Private
	#tag EndConstant

	#tag Constant, Name = UCM_NEEDPASSWORD, Type = Double, Dynamic = False, Default = \"2", Scope = Private
	#tag EndConstant

	#tag Constant, Name = UCM_NEEDPASSWORDW, Type = Double, Dynamic = False, Default = \"4", Scope = Private
	#tag EndConstant

	#tag Constant, Name = UCM_PROCESSDATA, Type = Double, Dynamic = False, Default = \"1", Scope = Private
	#tag EndConstant


	#tag ViewBehavior
		#tag ViewProperty
			Name="ArchivePassword"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
			InheritedFrom="UnRAR.Iterator"
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
