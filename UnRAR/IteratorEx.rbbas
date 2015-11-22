#tag Class
Protected Class IteratorEx
Inherits UnRAR.Iterator
	#tag Method, Flags = &h0
		Sub Constructor(RARFile As FolderItem, Mode As Integer, Password As String = "")
		  If Not UnRAR.IsAvailable Then Raise New PlatformNotSupportedException
		  mCommentBuffer = New MemoryBlock(260 * 2)
		  Dim path As New MemoryBlock(RARFile.AbsolutePath.LenB * 2 + 2)
		  Static UserData As Integer
		  UserData = UserData + 1
		  Dim UserDatum As Integer = UserData
		  If Instances = Nil Then Instances = New Dictionary
		  Instances.Value(UserDatum) = New WeakRef(Me)
		  path.WString(0) = RARFile.AbsolutePath
		  mArchiveHeader.CommentBufferSize = mCommentBuffer.Size
		  mArchiveHeader.Comments = mCommentBuffer
		  mArchiveHeader.ArchiveNameW = path
		  mArchiveHeader.Callback = AddressOf RARCallback
		  mArchiveHeader.UserData = UserDatum
		  mArchiveHeader.OpenMode = mode
		  ArchivePassword = Password
		  mHandle = RAROpenArchiveEx(mArchiveHeader)
		  mLastError = mArchiveHeader.OpenResult
		  If mHandle = 0 Then
		    Raise New RuntimeException
		  End If
		  
		  Dim h As RARHeaderDataEx
		  mLastError = RARReadHeaderEx(mHandle, h)
		  If mLastError <> 0 Then Raise New RuntimeException
		  mCurrentIndex = 0
		  mArchFile = RARFile
		  mCurrentItem = New UnRAR.ArchiveEntry(h, mCurrentIndex, mArchFile)
		End Sub
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
		    Dim header As RARHeaderDataEx
		    mLastError = RARReadHeaderEx(mHandle, header)
		    mCurrentIndex = mCurrentIndex + 1
		    If mLastError = 0 Then mCurrentItem = New UnRAR.ArchiveEntry(header, mCurrentIndex, mArchFile)
		  End If
		  
		  Return mLastError = 0
		End Function
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
		Sub Reset()
		  Me.Destructor
		  Me.Constructor(mArchFile, mOpenMode)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function _RARCallback(Message As UInt32, Param1 As Ptr, Param2 As Ptr) As Integer
		  Select Case Message
		  Case UCM_CHANGEVOLUME, UCM_CHANGEVOLUMEW
		    If Param2.UInt32(0) = RAR_VOL_ASK Then
		      Dim path As String = Param1.CString(0)
		      Dim f As FolderItem = GetFolderItem(path)
		      If RaiseEvent ChangeVolume(f) Then
		        Param1.CString(0) = f.AbsolutePath + Chr(0)
		        Return 1
		      End If
		    ElseIf Param2.UInt32(0) = RAR_VOL_NOTIFY Then
		      Return 1
		    End If
		    
		  Case UCM_PROCESSDATA
		    If Not RaiseEvent ProcessData(Param1, Integer(Param2)) Then Return 1
		    Return -1 ' abort
		    
		  Case UCM_NEEDPASSWORD
		    Dim mb As MemoryBlock = Param1
		    If ArchivePassword.Trim = "" And Not RaiseEvent PasswordRequired(ArchivePassword) Then Return -1 ' cancel
		    mb.CString(0) = ArchivePassword
		    Param2 = Ptr(ArchivePassword.Len)
		    Return 1
		    
		    'Case UCM_NEEDPASSWORDW
		    'Param1.WString(0) = ArchivePassword
		    'Return 1
		  End Select
		End Function
	#tag EndMethod


	#tag Hook, Flags = &h0
		Event ChangeVolume(ByRef NextVolume As FolderItem) As Boolean
	#tag EndHook

	#tag Hook, Flags = &h0
		Event PasswordRequired(ByRef Password As String) As Boolean
	#tag EndHook

	#tag Hook, Flags = &h0
		Event ProcessData(NewData As MemoryBlock, Length As Integer) As Boolean
	#tag EndHook


	#tag Property, Flags = &h21
		Private Shared Instances As Dictionary
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mArchiveHeader As RAROpenArchiveDataEx
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
