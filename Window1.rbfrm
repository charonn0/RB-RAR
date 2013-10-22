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
      Height          =   373
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
End
#tag EndWindow

#tag WindowCode
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
		    Return True
		  End If
		End Function
	#tag EndEvent
	#tag Event
		Function ContextualMenuAction(hitItem as MenuItem) As Boolean
		  Select Case Left(hitItem.Text, 7)
		  Case "Extract"
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
		    
		  End Select
		  Return True
		End Function
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
