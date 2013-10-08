#tag Module
Protected Module UnRAR
	#tag Method, Flags = &h0
		Function FormatError(RARErrorNumber As Integer) As String
		  Select Case RARErrorNumber
		  Case ErrorBadArchive
		    Return "Volume is not valid RAR archive"
		  Case ErrorBadData
		    Return "File CRC error"
		  Case ErrorEClose
		    Return "File close error"
		  Case ErrorECreate
		    Return "Output file create error"
		  Case ErrorEndArchive
		    Return "End of archive"
		  Case ErrorEOpen
		    Return "Volume open error"
		  Case ErrorERead
		    Return "Volume read error"
		  Case ErrorEWrite
		    Return "Output file write error"
		  Case ErrorNoMemory
		    Return "Insufficient memory for requested operation"
		  Case ErrorSmallBuff
		    Return "Output buffer is too small"
		  Case ErrorUnknown
		    Return "Unknown RAR error"
		  Case ErrorUnknownFormat
		    Return "Unknown archive format"
		  Else
		    Return "Unknown RAR error"
		  End Select
		  
		End Function
	#tag EndMethod

	#tag ExternalMethod, Flags = &h0
		Soft Declare Function RARCloseArchive Lib "UnRAR" (Handle As Integer) As Integer
	#tag EndExternalMethod

	#tag ExternalMethod, Flags = &h0
		Soft Declare Function RARGetDllVersion Lib "UnRAR" () As Integer
	#tag EndExternalMethod

	#tag ExternalMethod, Flags = &h0
		Soft Declare Function RAROpenArchive Lib "UnRAR" (ByRef Data As RAROpenArchiveData) As Integer
	#tag EndExternalMethod

	#tag ExternalMethod, Flags = &h0
		Soft Declare Function RARProcessFile Lib "UnRAR" (Handle As Integer, Operation As Integer, Desintation As CString, DestName As CString) As Integer
	#tag EndExternalMethod

	#tag ExternalMethod, Flags = &h0
		Soft Declare Function RARReadHeader Lib "UnRAR" (Handle As Integer, ByRef HeaderData As RARHeaderData) As Integer
	#tag EndExternalMethod

	#tag ExternalMethod, Flags = &h0
		Soft Declare Sub RARSetCallback Lib "UnRAR" (Handle As Integer, Callback As Ptr)
	#tag EndExternalMethod

	#tag ExternalMethod, Flags = &h0
		Soft Declare Sub RARSetPassword Lib "UnRAR" (Handle As Integer, Password As CString)
	#tag EndExternalMethod


	#tag Constant, Name = ErrorBadArchive, Type = Double, Dynamic = False, Default = \"13", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = ErrorBadData, Type = Double, Dynamic = False, Default = \"12", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = ErrorEClose, Type = Double, Dynamic = False, Default = \"17", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = ErrorECreate, Type = Double, Dynamic = False, Default = \"16", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = ErrorEndArchive, Type = Double, Dynamic = False, Default = \"10", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = ErrorEOpen, Type = Double, Dynamic = False, Default = \"15", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = ErrorERead, Type = Double, Dynamic = False, Default = \"18", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = ErrorEWrite, Type = Double, Dynamic = False, Default = \"19", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = ErrorNeedPassword, Type = Double, Dynamic = False, Default = \"255", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = ErrorNoMemory, Type = Double, Dynamic = False, Default = \"11", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = ErrorSmallBuff, Type = Double, Dynamic = False, Default = \"20", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = ErrorUnknown, Type = Double, Dynamic = False, Default = \"21", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = ErrorUnknownFormat, Type = Double, Dynamic = False, Default = \"14", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = OpTest, Type = Double, Dynamic = False, Default = \"1", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = RAR_DLL_VERSION, Type = Double, Dynamic = False, Default = \"3", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = RAR_EXTRACT, Type = Double, Dynamic = False, Default = \"2", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = RAR_OM_EXTRACT, Type = Double, Dynamic = False, Default = \"1", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = RAR_OM_LIST, Type = Double, Dynamic = False, Default = \"0", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = RAR_SKIP, Type = Double, Dynamic = False, Default = \"0", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = RAR_TEST, Type = Double, Dynamic = False, Default = \"1", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = RAR_VOL_ASK, Type = Double, Dynamic = False, Default = \"0", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = RAR_VOL_NOTIFY, Type = Double, Dynamic = False, Default = \"1", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = UCM_CHANGEVOLUME, Type = Double, Dynamic = False, Default = \"0", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = UCM_NEEDPASSWORD, Type = Double, Dynamic = False, Default = \"2", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = UCM_PROCESSDATA, Type = Double, Dynamic = False, Default = \"1", Scope = Protected
	#tag EndConstant


	#tag Structure, Name = RARHeaderData, Flags = &h0
		ArchiveName As String*260
		  FileName As String*260
		  Flags As UInt32
		  PackedSize As UInt32
		  UnpackedSize As UInt32
		  HostOS As UInt32
		  FileCRC As UInt32
		  FileTime As UInt32
		  UnpVer As UInt32
		  Method As UInt32
		  FileAttr As UInt32
		  CommentBuffer As Ptr
		  CommentBufferSize As UInt32
		  CommentSize As UInt32
		CommentState As UInt32
	#tag EndStructure

	#tag Structure, Name = RAROpenArchiveData, Flags = &h0
		AchiveName As Ptr
		  OpenMode As UInt32
		  OpenResult As UInt32
		  Comments As Ptr
		  CommentBufferSize As UInt32
		  CommentSize As UInt32
		CommentState As UInt32
	#tag EndStructure


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
End Module
#tag EndModule
