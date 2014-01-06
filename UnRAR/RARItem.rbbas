#tag Class
Class RARItem
	#tag Method, Flags = &h0
		Function Comment() As String
		  If RawData.StringValue(TargetLittleEndian) <> "" Then
		    If RawData.CommentState = 1 Then
		      Dim mb As MemoryBlock = RawData.CommentBuffer
		      Return mb.StringValue(0, RawData.CommentSize)
		    End If
		    
		  ElseIf RawDataEx.StringValue(TargetLittleEndian) <> "" Then
		    If RawDataEx.CommentState = 1 Then
		      Dim mb As MemoryBlock = RawDataEx.CommentBuffer
		      Return mb.StringValue(0, RawDataEx.CommentSize)
		    End If
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(Data As RARHeaderData, Index As Integer, RAR As FolderItem)
		  RawData = Data
		  RawDataEx.StringValue(TargetLittleEndian) = ""
		  mIndex = Index
		  mRARFile = RAR
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(Data As RARHeaderDataEx, Index As Integer, RAR As FolderItem)
		  RawDataEx = Data
		  RawData.StringValue(TargetLittleEndian) = ""
		  mIndex = Index
		  mRARFile = RAR
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function CRC32() As Integer
		  If RawData.StringValue(TargetLittleEndian) <> "" Then
		    Return RawData.FileCRC
		  ElseIf RawDataEx.StringValue(TargetLittleEndian) <> "" Then
		    Return RawDataEx.FileCRC
		  End If
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Directory() As Boolean
		  If RawData.StringValue(TargetLittleEndian) <> "" Then
		    Return BitAnd(Me.Flags, ItemFlag_Directory) = ItemFlag_Directory
		  ElseIf RawDataEx.StringValue(TargetLittleEndian) <> "" Then
		    Return BitAnd(Me.Flags, ItemFlag_Directory) = ItemFlag_Directory
		  End If
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function FileAttributes() As Integer
		  If RawData.StringValue(TargetLittleEndian) <> "" Then
		    Return RawData.FileAttr
		  ElseIf RawDataEx.StringValue(TargetLittleEndian) <> "" Then
		    Return RawDataEx.FileAttr
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function FileName() As String
		  If RawData.StringValue(TargetLittleEndian) <> "" Then
		    Return RawData.FileName
		  ElseIf RawDataEx.StringValue(TargetLittleEndian) <> "" Then
		    Return RawDataEx.FileName
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function FileTime() As Date
		  ' returns the modification date and time of the item as a Date object. Resolution is 1 second
		  Dim h, m, s, dom, mon, year As Integer
		  Dim dt, tm As UInt16
		  
		  If RawData.StringValue(TargetLittleEndian) <> "" Then
		    tm = RawData.FileTime
		    dt = RawData.FileDate
		  ElseIf RawDataEx.StringValue(TargetLittleEndian) <> "" Then
		    tm = RawDataEx.FileTime
		    dt = RawDataEx.FileDate
		  End If
		  
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
		  If RawData.StringValue(TargetLittleEndian) <> "" Then
		    Return RawData.Flags
		  ElseIf RawDataEx.StringValue(TargetLittleEndian) <> "" Then
		    Return RawDataEx.Flags
		  End If
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function HostOS() As String
		  Dim os As UInt32
		  If RawData.StringValue(TargetLittleEndian) <> "" Then
		    os = RawData.HostOS
		  ElseIf RawDataEx.StringValue(TargetLittleEndian) <> "" Then
		    os = RawDataEx.HostOS
		  End If
		  
		  Select Case os
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
		  If RawData.StringValue(TargetLittleEndian) <> "" Then
		    Return BitAnd(Me.Flags, ItemFlag_Encrypted) = ItemFlag_Encrypted
		  ElseIf RawDataEx.StringValue(TargetLittleEndian) <> "" Then
		    Return BitAnd(Me.Flags, ItemFlag_Encrypted) = ItemFlag_Encrypted
		  End If
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function IsSolid() As Boolean
		  If RawData.StringValue(TargetLittleEndian) <> "" Then
		    Return BitAnd(Me.Flags, ItemFlag_Solid) = ItemFlag_Solid
		  ElseIf RawDataEx.StringValue(TargetLittleEndian) <> "" Then
		    Return BitAnd(Me.Flags, ItemFlag_Solid) = ItemFlag_Solid
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function MinimumVersion() As Double
		  If RawData.StringValue(TargetLittleEndian) <> "" Then
		    Return RawData.UnpVer / 10
		  ElseIf RawDataEx.StringValue(TargetLittleEndian) <> "" Then
		    Return RawDataEx.UnpVer / 10
		  End If
		  
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
		  If RawData.StringValue(TargetLittleEndian) <> "" Then
		    Return RawData.PackedSize
		  ElseIf RawDataEx.StringValue(TargetLittleEndian) <> "" Then
		    Return RawDataEx.PackedSize
		  End If
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function PackingMethod() As Integer
		  If RawData.StringValue(TargetLittleEndian) <> "" Then
		    Return RawData.Method
		  ElseIf RawDataEx.StringValue(TargetLittleEndian) <> "" Then
		    Return RawDataEx.Method
		  End If
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function RARFile() As FolderItem
		  Return mRARFile
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function UnpackedSize() As UInt32
		  If RawData.StringValue(TargetLittleEndian) <> "" Then
		    Return RawData.UnpackedSize
		  ElseIf RawDataEx.StringValue(TargetLittleEndian) <> "" Then
		    Return RawDataEx.UnpackedSize
		  End If
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function VolumeName() As String
		  If RawData.StringValue(TargetLittleEndian) <> "" Then
		    Return RawData.ArchiveName
		  ElseIf RawDataEx.StringValue(TargetLittleEndian) <> "" Then
		    Return RawDataEx.ArchiveNameW
		  End If
		  
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

	#tag Property, Flags = &h1
		Protected RawDataEx As RARHeaderDataEx
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
