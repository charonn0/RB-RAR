#tag Module
Protected Module UnRAR
	#tag Method, Flags = &h0
		Function AbsolutePath_(Extends f As FolderItem) As String
		  #If RBVersion > 2019 Then
		    Return f.NativePath
		  #Else
		    Return f.AbsolutePath
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub ExtractArchive(RARFile As FolderItem, Destination As FolderItem, Password As String = "", ArchiveIndex As Integer = - 1)
		  If Not UnRAR.IsAvailable Then Raise New PlatformNotSupportedException
		  If Not RARFile.IsRARArchive Then Raise New UnsupportedFormatException
		  
		  Dim Archive As New UnRAR.IteratorEx(RARFile, RAR_OM_EXTRACT, Password)
		  Do
		    Select Case True
		    Case ArchiveIndex = -1, Archive.CurrentItem.Index = ArchiveIndex
		      Call Archive.MoveNext(UnRAR.RAR_EXTRACT, Destination)
		    Case Archive.CurrentItem.Index < ArchiveIndex
		      Call Archive.MoveNext(UnRAR.RAR_SKIP)
		    Else
		      Raise New OutOfBoundsException
		    End Select
		  Loop Until Archive.LastError <> 0
		  If Archive.LastError <> 0 And (Archive.LastError <> ErrorEndArchive Or ArchiveIndex > Archive.CurrentItem.Index) Then 
		    Raise New RARException(Archive.LastError)
		  End If
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function FormatError(RARErrorNumber As Integer) As String
		  Select Case RARErrorNumber
		  Case ErrorBadArchive
		    Return "The archive is corrupt."
		  Case ErrorBadData
		    Return "The archive is encrypted or corrupt."
		  Case ErrorEClose
		    Return "The archive can not be closed."
		  Case ErrorECreate
		    Return "The output file can not be created."
		  Case ErrorEndArchive
		    Return "The archive contains no additional entries."
		  Case ErrorEOpen
		    Return "The archive can not be opened."
		  Case ErrorERead
		    Return "The archive can not be read."
		  Case ErrorEWrite
		    Return "The output file can not be written."
		  Case ErrorNoMemory
		    Return "Out of memory!"
		  Case ErrorSmallBuff
		    Return "The output buffer is too small to contain the requested data."
		  Case ErrorNeedPassword
		    Return "A password is required in order to access the archive."
		  Case ErrorUnknown
		    Return "Unknown RAR error"
		  Case ErrorUnknownFormat
		    Return "The archive is of an unknown format."
		  Case ErrorRARUnavailable
		    Return "The UnRAR library is not installed."
		  Case ErrorUserCancel
		    Return "The operation was canceled."
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

	#tag Method, Flags = &h0
		Function IsRARArchive(Extends Arch As FolderItem) As Boolean
		  ' Returns True if the file is probably a RAR archive.
		  
		  Dim bs As BinaryStream = BinaryStream.Open(Arch)
		  Dim israr As Boolean
		  Try
		    israr = (bs.Read(4) = "Rar!")
		    If Not israr Then
		      bs.Position = 0
		      If bs.Read(2) = "MZ" Then ' an EXE file, maybe a RAR SFX
		        ' read the first 1MB+256 of the file and grep for the signature. The rar documentation states:
		        ' "At the moment of writing this documentation RAR assumes the maximum SFX module size to not exceed 1 MB"
		        Dim tmp As MemoryBlock = bs.Read(1024 * 1024 + 256)
		        israr = (InStr(tmp, "Rar!") > 0)
		      End If
		    End If
		  Finally
		    bs.Close
		  End Try
		  Return israr
		  
		Exception err As RuntimeException
		  If err IsA ThreadEndException Or err IsA EndException Then Raise err
		  Return False
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
		  If System.IsFunctionAvailable("RARGetDllVersion", "UnRAR") Then
		    Return RARGetDllVersion()
		  End If
		  Return -1
		End Function
	#tag EndMethod


	#tag Note, Name = Copying
		RB-RAR (https://github.com/charonn0/RB-RAR)
		Copyright (c)2013-23 Andrew Lambert
		
		Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated 
		documentation files (the "Software"), to deal in the Software without restriction, including without limitation 
		the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, 
		and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
		
		  The above copyright notice and this permission notice shall be included in all copies or substantial portions 
		  of the Software.
		
		THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO 
		THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE 
		AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, 
		TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
		
	#tag EndNote


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

	#tag Constant, Name = ErrorNeedPassword, Type = Double, Dynamic = False, Default = \"22", Scope = Protected
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

	#tag Constant, Name = ErrorUserCancel, Type = Double, Dynamic = False, Default = \"-1", Scope = Protected
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


	#tag Structure, Name = RARHeaderData, Flags = &h1
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

	#tag Structure, Name = RARHeaderDataEx, Flags = &h1
		ArchiveName As CString*1024
		  ArchiveNameW As WString*2048
		  FileName As CString*1024
		  FileNameW As WString*2048
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
		  UserData As Integer
		Reserved As Ptr
	#tag EndStructure


	#tag Enum, Name = PackingMethods, Flags = &h1
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
