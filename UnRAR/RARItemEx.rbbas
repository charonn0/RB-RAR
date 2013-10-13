#tag Class
Class RARItemEx
Inherits UnRAR.RARItem
	#tag Method, Flags = &h0
		Function Comment() As String
		  If RawData.CommentState = 1 Then
		    Dim mb As MemoryBlock = RawDataEx.CommentBuffer
		    Return mb.StringValue(0, RawDataEx.CommentSize)
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(Data As RARHeaderDataEx, Index As Integer)
		  RawDataEx = Data
		  mIndex = Index
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function CRC32() As Integer
		  Return RawDataEx.FileCRC
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Directory() As Boolean
		  Return BitAnd(Me.Flags, Flag_Directory) = Flag_Directory
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function FileAttributes() As Integer
		  Return RawDataEx.FileAttr
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function FileName() As String
		  Return RawDataEx.FileName
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function FileTime() As Date
		  ' returns the modification date and time of the item as a Date object. Resolution is 1 second
		  Dim h, m, s, dom, mon, year As Integer
		  Dim dt, tm As UInt16
		  tm = RawDataEx.FileTime
		  dt = RawDataEx.FileDate
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
		  Return RawDataEx.Flags
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function HostOS() As String
		  Select Case RawDataEx.HostOS
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
		Function IsEncrypted() As Boolean
		  Return BitAnd(Me.Flags, Flag_Encrypted) = Flag_Encrypted
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function IsSolid() As Boolean
		  Return BitAnd(Me.Flags, Flag_Solid) = Flag_Solid
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function MinimumVersion() As Double
		  Return RawDataEx.UnpVer / 10
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function PackedSize() As UInt32
		  Return RawDataEx.PackedSize
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function PackingMethod() As Integer
		  Return RawDataEx.Method
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function UnpackedSize() As UInt32
		  Return RawDataEx.UnpackedSize
		End Function
	#tag EndMethod


	#tag Property, Flags = &h1
		Protected RawDataEx As RARHeaderDataEx
	#tag EndProperty


End Class
#tag EndClass
