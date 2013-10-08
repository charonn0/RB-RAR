#tag Class
Class RARItem
	#tag Method, Flags = &h0
		Function ArchiveName() As String
		  Return RawData.ArchiveName
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Comment() As String
		  If RawData.CommentState = 1 Then
		    Dim mb As MemoryBlock = RawData.CommentBuffer
		    Return mb.StringValue(0, RawData.CommentSize)
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(Data As RARHeaderData, Index As Integer)
		  RawData = Data
		  mIndex = Index
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function CRC32() As Integer
		  Return RawData.FileCRC
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Encrypted() As Boolean
		  Return BitAnd(Me.Flags, Flag_Encrypted) = Flag_Encrypted
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function FileAttributes() As UInt32
		  Return RawData.FileAttr
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function FileName() As String
		  Return RawData.FileName
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function FileTime() As Date
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Flags() As UInt32
		  Return RawData.Flags
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function HostOS() As String
		  Select Case RawData.HostOS
		  Case 0
		    Return "MS-DOS"
		  Case 1
		    Return "OS/2"
		  Case 2
		    Return "Win32"
		  Case 3
		    Return "UNIX"
		  Else
		    Return "Unknown OS"
		  End Select
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Index() As Integer
		  Return mIndex
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function IsSolid() As Boolean
		  Return BitAnd(Me.Flags, Flag_Solid) = Flag_Solid
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function MinimumVersion() As Double
		  Return RawData.UnpVer / 10
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function PackedSize() As UInt32
		  Return RawData.PackedSize
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function PackingMethod() As UInt32
		  Return RawData.Method
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function UnpackedSize() As UInt32
		  Return RawData.UnpackedSize
		End Function
	#tag EndMethod


	#tag Property, Flags = &h1
		Protected mIndex As Integer
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected RawData As RARHeaderData
	#tag EndProperty


	#tag Constant, Name = Flag_CommentPresent, Type = Double, Dynamic = False, Default = \"&h08", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = Flag_ContinuedNext, Type = Double, Dynamic = False, Default = \"&h02", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = Flag_ContinuedPrev, Type = Double, Dynamic = False, Default = \"&h01", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = Flag_Encrypted, Type = Double, Dynamic = False, Default = \"&h04", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = Flag_Solid, Type = Double, Dynamic = False, Default = \"&h10", Scope = Protected
	#tag EndConstant


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
