#tag Module
Protected Module UnRAR
	#tag Method, Flags = &h21
		Private Sub CloseArchive(RARHandle As Integer)
		  If UnRAR.IsAvailable And RARHandle > 0 Then
		    Call RARCloseArchive(RARHandle)
		  End If
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function CRC32(Extends source As MemoryBlock) As Integer
		  Const CRCpolynomial = &hEDB88320
		  Dim crc, t as Integer
		  Dim strCode as String
		  strCode = source.StringValue(0, source.Size)
		  crc = &hffffffff
		  Dim char As String
		  
		  For x As Integer = 1 To LenB(strcode)
		    char = Midb(strcode, x, 1)
		    t = (crc And &hFF) Xor AscB(char)
		    For b As Integer = 0 To 7
		      If((t And &h1) = &h1) Then
		        t = bitwise.ShiftRight(t, 1, 32) Xor CRCpolynomial
		      Else
		        t = bitwise.ShiftRight(t, 1, 32)
		      End If
		    next
		    crc = Bitwise.ShiftRight(crc, 8, 32) Xor t
		  Next
		  Return (crc Xor &hFFFFFFFF)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function FormatError(RARErrorNumber As Integer) As String
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
		  Case ErrorRARUnavailable
		    Return "UnRAR library not available"
		  Case ErrorUserCancel
		    Return "Operation canceled."
		  Else
		    Return "Unknown RAR error: " + Str(RARErrorNumber)
		  End Select
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function IsAvailable() As Boolean
		  Return System.IsFunctionAvailable("RARProcessFile", "UnRAR")
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function IsAvailableEx() As Boolean
		  Return System.IsFunctionAvailable("RARProcessFileW", "UnRAR")
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function IsRARArchive(Extends Arch As FolderItem) As Boolean
		  Dim bs As BinaryStream = BinaryStream.Open(Arch)
		  Dim israr As Boolean = (bs.Read(4) = "Rar!")
		  bs.Close
		  Return israr
		  
		Exception err As RuntimeException
		  If err IsA ThreadEndException Or err IsA EndException Then Raise err
		  Return False
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function IterateArchive(RARFile As FolderItem, Mode As Integer = RAR_OM_EXTRACT, Password As String = "") As UnRAR.ArchiveIterator
		  Dim h As Integer = OpenArchive(RARFile, Mode)
		  If h > 0 Then
		    Return New UnRAR.ArchiveIterator(h, RARFile, Mode, Password)
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function IterateArchiveEx(RARFile As FolderItem, Mode As Integer = RAR_OM_EXTRACT, Password As String = "") As UnRAR.ArchiveIteratorEx
		  Return New UnRAR.ArchiveIteratorEx(RARFile, Mode, Password)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function OpenArchive(RARFile As FolderItem, Mode As Integer) As Integer
		  If UnRAR.IsAvailable Then
		    Dim mHandle, err As Integer
		    Dim mb As New MemoryBlock(260 * 2)
		    Dim path As New MemoryBlock(RARFile.AbsolutePath.LenB * 2)
		    path.CString(0) = RARFile.AbsolutePath
		    Dim data As RAROpenArchiveData
		    data.CommentBufferSize = mb.Size
		    data.Comments = mb
		    data.ArchiveName = path
		    data.OpenMode = mode
		    mHandle = UnRAR.RAROpenArchive(data)
		    err = data.OpenResult
		    If mHandle <= 0 Then Return err * -1
		    Return mHandle
		  End If
		End Function
	#tag EndMethod

	#tag ExternalMethod, Flags = &h21
		Private Soft Declare Function RARCloseArchive Lib "UnRAR" (Handle As Integer) As Integer
	#tag EndExternalMethod

	#tag ExternalMethod, Flags = &h21
		Private Soft Declare Function RARGetDllVersion Lib "UnRAR" () As Integer
	#tag EndExternalMethod

	#tag ExternalMethod, Flags = &h21
		Private Soft Declare Function RAROpenArchive Lib "UnRAR" (ByRef Data As RAROpenArchiveData) As Integer
	#tag EndExternalMethod

	#tag ExternalMethod, Flags = &h21
		Private Soft Declare Function RAROpenArchiveEx Lib "UnRAR" (ByRef Data As RAROpenArchiveDataEx) As Integer
	#tag EndExternalMethod

	#tag ExternalMethod, Flags = &h21
		Private Soft Declare Function RARProcessFile Lib "UnRAR" (Handle As Integer, Operation As Integer, Desintation As CString, DestName As CString) As Integer
	#tag EndExternalMethod

	#tag ExternalMethod, Flags = &h21
		Private Soft Declare Function RARProcessFileW Lib "UnRAR" (Handle As Integer, Operation As Integer, Desintation As CString, DestName As CString) As Integer
	#tag EndExternalMethod

	#tag ExternalMethod, Flags = &h21
		Private Soft Declare Function RARReadHeader Lib "UnRAR" (Handle As Integer, ByRef HeaderData As RARHeaderData) As Integer
	#tag EndExternalMethod

	#tag ExternalMethod, Flags = &h21
		Private Soft Declare Function RARReadHeaderEx Lib "UnRAR" (Handle As Integer, ByRef HeaderData As RARHeaderDataEx) As Integer
	#tag EndExternalMethod

	#tag ExternalMethod, Flags = &h21
		Private Soft Declare Sub RARSetPassword Lib "UnRAR" (Handle As Integer, Password As CString)
	#tag EndExternalMethod

	#tag Method, Flags = &h1
		Protected Function Version() As Integer
		  If UnRAR.IsAvailable Then
		    Return RARGetDllVersion()
		  End If
		End Function
	#tag EndMethod


	#tag Note, Name = Copying
		The source code in this project is released under the terms of the MIT License, 
		reproduced below. In addition to the terms of the license, it is requested that 
		anyone with suggestions, bugs, feature requests, or their own code to contribute 
		contact me at andrew[at]boredomsoft[dot]org
		
		--------------------------------------------------------------------------------
		Copyright (c) 2013-14 Andrew Lambert
		
		Permission is hereby granted, free of charge, to any person obtaining a copy
		of this software and associated documentation files (the "Software"), to deal
		in the Software without restriction, including without limitation the rights
		to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
		copies of the Software, and to permit persons to whom the Software is
		furnished to do so, subject to the following conditions:
		
		The above copyright notice and this permission notice shall be included in
		all copies or substantial portions of the Software.
		
		THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
		IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
		FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
		AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
		LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
		OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
		THE SOFTWARE.
	#tag EndNote


	#tag Constant, Name = ArchiveFlag_AuthenticityInfo, Type = Double, Dynamic = False, Default = \"&h020", Scope = Private
	#tag EndConstant

	#tag Constant, Name = ArchiveFlag_Comment, Type = Double, Dynamic = False, Default = \"&h002", Scope = Private
	#tag EndConstant

	#tag Constant, Name = ArchiveFlag_EncryptedNames, Type = Double, Dynamic = False, Default = \"&h080", Scope = Private
	#tag EndConstant

	#tag Constant, Name = ArchiveFlag_FirstVolume, Type = Double, Dynamic = False, Default = \"&h100", Scope = Private
	#tag EndConstant

	#tag Constant, Name = ArchiveFlag_Locked, Type = Double, Dynamic = False, Default = \"&h004", Scope = Private
	#tag EndConstant

	#tag Constant, Name = ArchiveFlag_NewNamingScheme, Type = Double, Dynamic = False, Default = \"&h010", Scope = Private
	#tag EndConstant

	#tag Constant, Name = ArchiveFlag_RecoveryRecord, Type = Double, Dynamic = False, Default = \"&h040", Scope = Private
	#tag EndConstant

	#tag Constant, Name = ArchiveFlag_Solid, Type = Double, Dynamic = False, Default = \"&h008", Scope = Private
	#tag EndConstant

	#tag Constant, Name = ArchiveFlag_Volume, Type = Double, Dynamic = False, Default = \"&h001", Scope = Private
	#tag EndConstant

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

	#tag Constant, Name = ErrorRARUnavailable, Type = Double, Dynamic = False, Default = \"254", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = ErrorSmallBuff, Type = Double, Dynamic = False, Default = \"20", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = ErrorUnknown, Type = Double, Dynamic = False, Default = \"21", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = ErrorUnknownFormat, Type = Double, Dynamic = False, Default = \"14", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = ErrorUserCancel, Type = Double, Dynamic = False, Default = \"-1", Scope = Public
	#tag EndConstant

	#tag Constant, Name = ItemFlag_CommentPresent, Type = Double, Dynamic = False, Default = \"&h08", Scope = Private
	#tag EndConstant

	#tag Constant, Name = ItemFlag_ContinuedNext, Type = Double, Dynamic = False, Default = \"&h02", Scope = Private
	#tag EndConstant

	#tag Constant, Name = ItemFlag_ContinuedPrev, Type = Double, Dynamic = False, Default = \"&h01", Scope = Private
	#tag EndConstant

	#tag Constant, Name = ItemFlag_Directory, Type = Double, Dynamic = False, Default = \"&h20", Scope = Private
	#tag EndConstant

	#tag Constant, Name = ItemFlag_Encrypted, Type = Double, Dynamic = False, Default = \"&h04", Scope = Private
	#tag EndConstant

	#tag Constant, Name = ItemFlag_Solid, Type = Double, Dynamic = False, Default = \"&h10", Scope = Private
	#tag EndConstant

	#tag Constant, Name = RAR_DLL_VERSION, Type = Double, Dynamic = False, Default = \"3", Scope = Private
	#tag EndConstant

	#tag Constant, Name = RAR_EXTRACT, Type = Double, Dynamic = False, Default = \"2", Scope = Private
	#tag EndConstant

	#tag Constant, Name = RAR_OM_EXTRACT, Type = Double, Dynamic = False, Default = \"1", Scope = Private
	#tag EndConstant

	#tag Constant, Name = RAR_OM_LIST, Type = Double, Dynamic = False, Default = \"0", Scope = Private
	#tag EndConstant

	#tag Constant, Name = RAR_SKIP, Type = Double, Dynamic = False, Default = \"0", Scope = Private
	#tag EndConstant

	#tag Constant, Name = RAR_TEST, Type = Double, Dynamic = False, Default = \"1", Scope = Private
	#tag EndConstant

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


	#tag Structure, Name = RARHeaderData, Flags = &h21
		ArchiveName As CString*260
		  FileName As CString*260
		  Flags As UInt32
		  PackedSize As UInt32
		  UnpackedSize As UInt32
		  HostOS As UInt32
		  FileCRC As UInt32
		  FileTime As UInt16
		  FileDate As UInt16
		  UnpVer As UInt32
		  Method As UInt32
		  FileAttr As UInt32
		  CommentBuffer As Ptr
		  CommentBufferSize As UInt32
		  CommentSize As UInt32
		CommentState As UInt32
	#tag EndStructure

	#tag Structure, Name = RARHeaderDataEx, Flags = &h21
		ArchiveName As String*1024
		  ArchiveNameW As WString*1024
		  FileName As CString*1024
		  FileNameW As WString*1024
		  Flags As UInt32
		  PackedSize As UInt32
		  PackedSizeHi As UInt32
		  UnpackedSize As UInt32
		  UnpackedSizeHi As UInt32
		  HostOS As UInt32
		  FileCRC As UInt32
		  FileTime As UInt16
		  FileDate As UInt16
		  UnpVer As UInt32
		  Method As UInt32
		  FileAttr As UInt32
		  CommentBuffer As Ptr
		  CommentBufferSize As UInt32
		  CommentSize As UInt32
		  CommentState As UInt32
		  DictSize As UInt32
		  HashType As UInt32
		  Hash As String*32
		Reserved As String*1014
	#tag EndStructure

	#tag Structure, Name = RAROpenArchiveData, Flags = &h21
		ArchiveName As Ptr
		  OpenMode As UInt32
		  OpenResult As UInt32
		  Comments As Ptr
		  CommentBufferSize As UInt32
		  CommentSize As UInt32
		CommentState As UInt32
	#tag EndStructure

	#tag Structure, Name = RAROpenArchiveDataEx, Flags = &h21
		ArchiveName As Ptr
		  ArchiveNameW As Ptr
		  OpenMode As UInt32
		  OpenResult As UInt32
		  Comments As Ptr
		  CommentBufferSize As UInt32
		  CommentSize As UInt32
		  CommentState As UInt32
		  Flags As UInt32
		  Callback As Ptr
		  UserData As Ptr
		Reserved As Ptr
	#tag EndStructure


	#tag Enum, Name = PackingMethods, Type = Integer, Flags = &h1
		Store=&h30
		  Fastest=&h31
		  Fast=&h32
		  Normal=&h33
		  Good=&h34
		Best=&h35
	#tag EndEnum


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
