#tag Class
Class RARItem
	#tag Method, Flags = &h0
		Function Comment() As String
		  If RawData.CommentState = 1 Then
		    Dim mb As MemoryBlock = RawData.CommentBuffer
		    Return mb.StringValue(0, RawData.CommentSize)
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(Data As RARHeaderData, Index As Integer, RAR As FolderItem)
		  RawData = Data
		  mIndex = Index
		  mRARFile = RAR
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function CRC32() As Integer
		  Return RawData.FileCRC
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Directory() As Boolean
		  Return BitAnd(Me.Flags, ItemFlag_Directory) = ItemFlag_Directory
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function FileAttributes() As Integer
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
		  ' returns the modification date and time of the item as a Date object. Resolution is 1 second
		  Dim h, m, s, dom, mon, year As Integer
		  Dim dt, tm As UInt16
		  tm = RawData.FileTime
		  dt = RawData.FileDate
		  h = Bitwise.ShiftRight(tm, 11)
		  m = Bitwise.ShiftRight(tm, 5) And &h3F
		  s = (tm And &h1F) * 2
		  dom = dt And &h1F
		  mon = Bitwise.ShiftRight(dt, 5) And &h0F
		  year = (Bitwise.ShiftRight(dt, 9) And &h7F) + 1980
		  Return New Date(year, mon, dom, h, m, s)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Flags() As Integer
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
		Function IsEncrypted() As Boolean
		  Return BitAnd(Me.Flags, ItemFlag_Encrypted) = ItemFlag_Encrypted
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function IsSolid() As Boolean
		  Return BitAnd(Me.Flags, ItemFlag_Solid) = ItemFlag_Solid
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function MinimumVersion() As Double
		  Return RawData.UnpVer / 10
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Operator_Compare(OtherItem As RARItem) As Integer
		  If OtherItem Is Nil Then Return -1
		  If OtherItem.RARFile.AbsolutePath <> Me.RARFile.AbsolutePath Then Return -1
		  If OtherItem.Index <> Me.Index Then Return 1
		  Return 0
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function PackedSize() As UInt32
		  Return RawData.PackedSize
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function PackingMethod() As Integer
		  Return RawData.Method
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function RARFile() As FolderItem
		  Return mRARFile
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function UnpackedSize() As UInt32
		  Return RawData.UnpackedSize
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function VolumeName() As String
		  Return RawData.ArchiveName
		End Function
	#tag EndMethod


	#tag Property, Flags = &h1
		Protected mIndex As Integer
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mRARFile As FolderItem
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected RawData As RARHeaderData
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
