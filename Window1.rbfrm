#tag Window
Begin Window Window1
   BackColor       =   16777215
   Backdrop        =   ""
   CloseButton     =   True
   Composite       =   False
   Frame           =   0
   FullScreen      =   False
   HasBackColor    =   False
   Height          =   3.93e+2
   ImplicitInstance=   True
   LiveResize      =   True
   MacProcID       =   0
   MaxHeight       =   32000
   MaximizeButton  =   False
   MaxWidth        =   32000
   MenuBar         =   1054722047
   MenuBarVisible  =   True
   MinHeight       =   64
   MinimizeButton  =   True
   MinWidth        =   64
   Placement       =   0
   Resizeable      =   True
   Title           =   "Untitled"
   Visible         =   True
   Width           =   6.0e+2
   Begin Listbox Listbox1
      AutoDeactivate  =   True
      AutoHideScrollbars=   True
      Bold            =   ""
      Border          =   True
      ColumnCount     =   3
      ColumnsResizable=   ""
      ColumnWidths    =   "50%,25%,25%"
      DataField       =   ""
      DataSource      =   ""
      DefaultRowHeight=   -1
      Enabled         =   True
      EnableDrag      =   ""
      EnableDragReorder=   ""
      GridLinesHorizontal=   0
      GridLinesVertical=   0
      HasHeading      =   True
      HeadingIndex    =   -1
      Height          =   249
      HelpTag         =   ""
      Hierarchical    =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      InitialValue    =   "Name	Size	Packed"
      Italic          =   ""
      Left            =   0
      LockBottom      =   ""
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   ""
      LockTop         =   True
      RequiresSelection=   ""
      Scope           =   0
      ScrollbarHorizontal=   ""
      ScrollBarVertical=   True
      SelectionType   =   0
      TabIndex        =   0
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   0
      TextUnit        =   0
      Top             =   0
      Underline       =   ""
      UseFocusRing    =   True
      Visible         =   True
      Width           =   600
      _ScrollWidth    =   -1
   End
   Begin PushButton PushButton1
      AutoDeactivate  =   True
      Bold            =   ""
      ButtonStyle     =   0
      Cancel          =   ""
      Caption         =   "Open RAR"
      Default         =   ""
      Enabled         =   True
      Height          =   22
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   ""
      Left            =   269
      LockBottom      =   ""
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   ""
      LockTop         =   True
      Scope           =   0
      TabIndex        =   1
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   0
      TextUnit        =   0
      Top             =   371
      Underline       =   ""
      Visible         =   True
      Width           =   80
   End
   Begin Listbox Listbox2
      AutoDeactivate  =   True
      AutoHideScrollbars=   True
      Bold            =   ""
      Border          =   True
      ColumnCount     =   2
      ColumnsResizable=   ""
      ColumnWidths    =   ""
      DataField       =   ""
      DataSource      =   ""
      DefaultRowHeight=   -1
      Enabled         =   True
      EnableDrag      =   ""
      EnableDragReorder=   ""
      GridLinesHorizontal=   0
      GridLinesVertical=   0
      HasHeading      =   True
      HeadingIndex    =   -1
      Height          =   98
      HelpTag         =   ""
      Hierarchical    =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      InitialValue    =   "Property	Value"
      Italic          =   ""
      Left            =   0
      LockBottom      =   ""
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   ""
      LockTop         =   True
      RequiresSelection=   ""
      Scope           =   0
      ScrollbarHorizontal=   ""
      ScrollBarVertical=   True
      SelectionType   =   0
      TabIndex        =   2
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   0
      TextUnit        =   0
      Top             =   261
      Underline       =   ""
      UseFocusRing    =   True
      Visible         =   True
      Width           =   600
      _ScrollWidth    =   -1
   End
End
#tag EndWindow

#tag WindowCode
	#tag Event
		Sub Open()
		  MsgBox(Str(UnRAR.Version))
		End Sub
	#tag EndEvent


	#tag Property, Flags = &h1
		Protected Archive As UnRAR.RARchive
	#tag EndProperty


#tag EndWindowCode

#tag Events Listbox1
	#tag Event
		Function ConstructContextualMenu(base as MenuItem, x as Integer, y as Integer) As Boolean
		  Dim row As Integer = Me.RowFromXY(X, Y)
		  If row <= Me.ListCount - 1 Then
		    Dim item As RARItem = Me.RowTag(row)
		    Dim extract As New MenuItem("Extract " + item.FileName)
		    extract.Tag = item
		    base.Append(extract)
		    base.Append(New MenuItem("Extract All"))
		    Dim tst As New MenuItem("Test " + item.FileName)
		    tst.Tag = item
		    base.Append(tst)
		    base.Append(New MenuItem("Test All"))
		    Return True
		  End If
		End Function
	#tag EndEvent
	#tag Event
		Function ContextualMenuAction(hitItem as MenuItem) As Boolean
		  Select Case Left(hitItem.Text, 4)
		  Case "Extr"
		    Dim item As RARItem = hitItem.Tag
		    If item <> Nil Then
		      Dim f As FolderItem = GetSaveFolderItem("", item.FileName)
		      If f <> Nil Then
		        If Self.Archive.ExtractItem(item.Index, f) Then
		          MsgBox("Extracted")
		        Else
		          MsgBox("RAR error " + Str(Self.Archive.LastError) + ": " + UnRAR.FormatError(Self.Archive.LastError))
		        End If
		      End If
		    Else
		      Dim f As FolderItem = SelectFolder()
		      If f <> Nil Then
		        If Self.Archive.ExtractAll(f) Then
		          MsgBox("Extracted")
		        Else
		          MsgBox("RAR error " + Str(Self.Archive.LastError) + ": " + UnRAR.FormatError(Self.Archive.LastError))
		        End If
		      End If
		    End If
		    
		  Case "Test"
		    Dim item As RARItem = hitItem.Tag
		    If item <> Nil Then
		      If Self.Archive.TestItem(item.Index) Then
		        MsgBox("Test OK")
		      Else
		        MsgBox("RAR error " + Str(Self.Archive.LastError) + ": " + UnRAR.FormatError(Self.Archive.LastError))
		      End If
		    Else
		      If Self.Archive.TestAll() Then
		        MsgBox("Test OK")
		      Else
		        MsgBox("RAR error " + Str(Self.Archive.LastError) + ": " + UnRAR.FormatError(Self.Archive.LastError))
		      End If
		    End If
		    
		  End Select
		  Return True
		End Function
	#tag EndEvent
	#tag Event
		Sub Change()
		  Listbox2.DeleteAllRows
		  If Me.ListIndex > -1 Then
		    Dim item As RARItem = Me.RowTag(Me.ListIndex)
		    If item = Nil Then Return
		    Listbox2.AddRow("File Name",  item.FileName)
		    Listbox2.AddRow("Time",  item.FileTime.SQLDateTime)
		    Listbox2.AddRow("Archive", item.RARFile.AbsolutePath)
		    Listbox2.AddRow("Volume", item.VolumeName)
		    Listbox2.AddRow("Comment",  item.Comment)
		    Listbox2.AddRow("Packed size",  Format(item.PackedSize, "###,###,###,###"))
		    Listbox2.AddRow("Unpacked size",  Format(item.UnpackedSize, "###,###,###,###"))
		    Listbox2.AddRow("Index",  Str(item.Index))
		    Listbox2.AddRow("Encrypted",  Str(item.IsEncrypted))
		    Listbox2.AddRow("Solid",  Str(item.IsSolid))
		    Listbox2.AddRow("CRC32",  "0x" + Left(Hex(item.CRC32) + "00000000", 8))
		    Listbox2.AddRow("Directory",  Str(item.Directory))
		    Listbox2.AddRow("Attributes",  "0x" + Left(Hex(item.FileAttributes) + "00000000", 8))
		    Listbox2.AddRow("Flags",  "0x" + Left(Hex(item.Flags) + "00000000", 8))
		    Listbox2.AddRow("OS",  item.HostOS)
		    Listbox2.AddRow("Min. Version",  Format(item.MinimumVersion, "#0.0#"))
		    Listbox2.AddRow("Packing Method",  Str(item.PackingMethod))
		  End If
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events PushButton1
	#tag Event
		Sub Action()
		  Dim rar As FolderItem = GetOpenFolderItem(FileTypes1.All)
		  If rar <> Nil And rar.IsRARArchive Then
		    Listbox1.DeleteAllRows
		    Self.Archive = New RARchive(rar)
		    Dim count As Integer = Archive.Count - 1
		    For i As Integer = 0 To count
		      Dim item As RARItem = Archive.Item(i)
		      If item.Directory Then
		        Listbox1.AddFolder(item.FileName)
		      Else
		        Dim d, sz, p As String
		        d = item.FileTime.SQLDateTime
		        sz = Format(item.UnpackedSize, "###,###,###,###")
		        p = Format(item.PackedSize * 100 / item.UnpackedSize, "#0.0#\%")
		        Listbox1.AddRow(item.FileName, sz, p, d)
		      End If
		      Listbox1.RowTag(Listbox1.LastIndex) = item
		    Next
		  End If
		End Sub
	#tag EndEvent
#tag EndEvents
