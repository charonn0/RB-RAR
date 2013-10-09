#tag Class
Class RARItem
	#tag Method, Flags = &h0
		Sub Constructor(Data As RARHeaderData, Index As Integer)
		  RawData = Data
		  mIndex = Index
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Index() As Integer
		  Return mIndex
		End Function
	#tag EndMethod


	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  Return GetFolderItem(RawData.ArchiveName)
			End Get
		#tag EndGetter
		Archive As FolderItem
	#tag EndComputedProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  If RawData.CommentState = 1 Then
			    Dim mb As MemoryBlock = RawData.CommentBuffer
			    Return mb.StringValue(0, RawData.CommentSize)
			  End If
			End Get
		#tag EndGetter
		Comment As String
	#tag EndComputedProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  Return RawData.FileCRC
			End Get
		#tag EndGetter
		CRC32 As Integer
	#tag EndComputedProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  Return BitAnd(Me.Flags, Flag_Directory) = Flag_Directory
			End Get
		#tag EndGetter
		Directory As Boolean
	#tag EndComputedProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  Return RawData.FileAttr
			End Get
		#tag EndGetter
		FileAttributes As UInt32
	#tag EndComputedProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  Return RawData.FileName
			End Get
		#tag EndGetter
		FileName As String
	#tag EndComputedProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
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
			End Get
		#tag EndGetter
		FileTime As Date
	#tag EndComputedProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  Return RawData.Flags
			End Get
		#tag EndGetter
		Flags As UInt32
	#tag EndComputedProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
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
			End Get
		#tag EndGetter
		HostOS As String
	#tag EndComputedProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  Return BitAnd(Me.Flags, Flag_Encrypted) = Flag_Encrypted
			End Get
		#tag EndGetter
		IsEncrypted As Boolean
	#tag EndComputedProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  Return BitAnd(Me.Flags, Flag_Solid) = Flag_Solid
			End Get
		#tag EndGetter
		IsSolid As Boolean
	#tag EndComputedProperty

	#tag Property, Flags = &h1
		Protected mIndex As Integer
	#tag EndProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  Return RawData.UnpVer / 10
			End Get
		#tag EndGetter
		MinimumVersion As Double
	#tag EndComputedProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  Return RawData.PackedSize
			End Get
		#tag EndGetter
		PackedSize As UInt32
	#tag EndComputedProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  Return RawData.Method
			End Get
		#tag EndGetter
		PackingMethod As UInt32
	#tag EndComputedProperty

	#tag Property, Flags = &h1
		Protected RawData As RARHeaderData
	#tag EndProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  Return RawData.UnpackedSize
			End Get
		#tag EndGetter
		UnpackedSize As UInt32
	#tag EndComputedProperty


	#tag ViewBehavior
		#tag ViewProperty
			Name="Comment"
			Group="Behavior"
			Type="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="CRC32"
			Group="Behavior"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Directory"
			Group="Behavior"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="FileName"
			Group="Behavior"
			Type="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="HostOS"
			Group="Behavior"
			Type="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			InitialValue="-2147483648"
			InheritedFrom="Object"
		#tag EndViewProperty
		#tag ViewProperty
			Name="IsEncrypted"
			Group="Behavior"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="IsSolid"
			Group="Behavior"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=true
			Group="Position"
			InitialValue="0"
			InheritedFrom="Object"
		#tag EndViewProperty
		#tag ViewProperty
			Name="MinimumVersion"
			Group="Behavior"
			Type="Double"
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
