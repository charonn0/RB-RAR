#tag Class
Class RARItem
	#tag Method, Flags = &h0
		Function Comment() As String
		  If Not IsEx Then
		    If RawData.CommentState = 1 Then
		      Dim mb As MemoryBlock = RawData.CommentBuffer
		      Return mb.StringValue(0, RawData.CommentSize)
		    End If
		    
		  Else
		    If RawDataEx.CommentState = 1 Then
		      Dim mb As MemoryBlock = RawDataEx.CommentBuffer
		      Return mb.StringValue(0, RawDataEx.CommentSize)
		    End If
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(Data As RARHeaderData, Index As Integer, RAR As FolderItem)
		  IsEx = False
		  RawData = Data
		  RawDataEx.StringValue(TargetLittleEndian) = ""
		  mIndex = Index
		  mRARFile = RAR
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(Data As RARHeaderDataEx, Index As Integer, RAR As FolderItem)
		  IsEx = True
		  RawDataEx = Data
		  RawData.StringValue(TargetLittleEndian) = ""
		  mIndex = Index
		  mRARFile = RAR
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function CRC32() As Integer
		  If Not IsEx Then
		    Return RawData.FileCRC
		  Else
		    Return RawDataEx.FileCRC
		  End If
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Directory() As Boolean
		  If Not IsEx Then
		    Return BitAnd(Me.Flags, ItemFlag_Directory) = ItemFlag_Directory
		  Else
		    Return BitAnd(Me.Flags, ItemFlag_Directory) = ItemFlag_Directory
		  End If
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function FileAttributes() As Integer
		  If Not IsEx Then
		    Return RawData.FileAttr
		  Else
		    Return RawDataEx.FileAttr
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function FileName() As String
		  If Not IsEx Then
		    Return RawData.FileName
		  Else
		    Return RawDataEx.FileNameW
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function FileTime() As Date
		  ' returns the modification date and time of the item as a Date object.
		  ' WinRAR stores timestamps in MS-DOS format. Resolution is 2 seconds
		  Dim h, m, s, dom, mon, year As Integer
		  Dim dt, tm As UInt16
		  
		  If Not IsEx Then
		    tm = RawData.FileTime
		    dt = RawData.FileDate
		  Else
		    tm = RawDataEx.FileTime
		    dt = RawDataEx.FileDate
		  End If
		  h = ShiftRight(tm, 11)
		  m = ShiftRight(tm, 5) And &h3F
		  s = (tm And &h1F) * 2
		  dom = dt And &h1F
		  mon = ShiftRight(dt, 5) And &h0F
		  year = (ShiftRight(dt, 9) And &h7F) + 1980
		  
		  Return New Date(year, mon, dom, h, m, s)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Flags() As Integer
		  If Not IsEx Then
		    Return RawData.Flags
		  Else
		    Return RawDataEx.Flags
		  End If
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function HostOS() As String
		  Dim os As UInt32
		  If Not IsEx Then
		    os = RawData.HostOS
		  Else
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
		  If Not IsEx Then
		    Return RawData.UnpVer / 10
		  Else
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
		  If Not IsEx Then
		    Return RawData.PackedSize
		  Else
		    Return BitOr(ShiftLeft(RawDataEx.UnpackedSizeHi, 32), RawDataEx.UnpackedSize)
		  End If
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function PackingMethod() As Integer
		  If Not IsEx Then
		    Return RawData.Method
		  Else
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
		Function UnpackedSize() As UInt64
		  If Not IsEx Then
		    Return RawData.UnpackedSize
		  Else
		    Return BitOr(ShiftLeft(RawDataEx.UnpackedSizeHi, 32), RawDataEx.UnpackedSize)
		  End If
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function VolumeName() As String
		  If Not IsEx Then
		    Return RawData.ArchiveName
		  Else
		    Return DefineEncoding(RawDataEx.ArchiveNameW, Encodings.UTF16)
		  End If
		  
		End Function
	#tag EndMethod


	#tag Property, Flags = &h21
		Private IsEx As Boolean
	#tag EndProperty

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
