#tag Window
Begin Window Window1
   BackColor       =   16777215
   Backdrop        =   ""
   CloseButton     =   True
   Composite       =   False
   Frame           =   0
   FullScreen      =   False
   HasBackColor    =   False
   Height          =   5.85e+2
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
   Placement       =   2
   Resizeable      =   False
   Title           =   "UnRar demo"
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
      Height          =   258
      HelpTag         =   ""
      Hierarchical    =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      InitialValue    =   "Name	Size	Packed"
      Italic          =   ""
      Left            =   0
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
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
      Top             =   29
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
      Left            =   520
      LockBottom      =   ""
      LockedInPosition=   False
      LockLeft        =   False
      LockRight       =   True
      LockTop         =   True
      Scope           =   0
      TabIndex        =   1
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   0
      TextUnit        =   0
      Top             =   2
      Underline       =   ""
      Visible         =   True
      Width           =   80
   End
   Begin GroupBox GroupBox1
      AutoDeactivate  =   True
      Bold            =   ""
      Caption         =   "Selected Item Properties"
      Enabled         =   True
      Height          =   269
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   ""
      Left            =   6
      LockBottom      =   ""
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   ""
      LockTop         =   True
      Scope           =   0
      TabIndex        =   4
      TabPanelIndex   =   0
      TextFont        =   "System"
      TextSize        =   0
      TextUnit        =   0
      Top             =   307
      Underline       =   ""
      Visible         =   True
      Width           =   588
   End
   Begin Label Label1
      AutoDeactivate  =   True
      Bold            =   ""
      DataField       =   ""
      DataSource      =   ""
      Enabled         =   True
      Height          =   20
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   ""
      Left            =   6
      LockBottom      =   ""
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   True
      Multiline       =   ""
      Scope           =   0
      Selectable      =   False
      TabIndex        =   3
      TabPanelIndex   =   0
      Text            =   "No Archive"
      TextAlign       =   0
      TextColor       =   &h000000
      TextFont        =   "System"
      TextSize        =   0
      TextUnit        =   0
      Top             =   2
      Transparent     =   False
      Underline       =   ""
      Visible         =   True
      Width           =   510
   End
   Begin ProgressBar ProgressBar1
      AutoDeactivate  =   True
      Enabled         =   True
      Height          =   9
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Left            =   0
      LockBottom      =   ""
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   ""
      LockTop         =   True
      Maximum         =   100
      Scope           =   0
      TabPanelIndex   =   0
      Top             =   292
      Value           =   0
      Visible         =   True
      Width           =   600
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
      Height          =   237
      HelpTag         =   ""
      Hierarchical    =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      InitialValue    =   "Property	Value"
      Italic          =   ""
      Left            =   20
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   False
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
      Top             =   328
      Underline       =   ""
      UseFocusRing    =   True
      Visible         =   True
      Width           =   560
      _ScrollWidth    =   -1
   End
End
#tag EndWindow

#tag WindowCode
	#tag Method, Flags = &h0
		Sub ProcessHandler(Sender As RARchive, Item As RARItem, Operation As Integer, ByRef SaveTo As FolderItem)
		  #pragma Unused SaveTo
		  #pragma Unused Sender
		  System.DebugLog(item.FileName + " processed in mode: " + Str(Operation))
		  ProgressBar1.Value = Item.Index
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h1
		Protected Archive As UnRAR.RARchive
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected LastTop As Integer
	#tag EndProperty


#tag EndWindowCode

#tag Events Listbox1
	#tag Event
		Function ConstructContextualMenu(base as MenuItem, x as Integer, y as Integer) As Boolean
		  Dim row As Integer = Me.RowFromXY(X, Y)
		  If row > -1 Then
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
		  Dim item As RARItem = hitItem.Tag
		  Dim pw, msg As String
		  Dim index As Integer
		  
		  If item <> Nil And item.IsEncrypted Then
		    pw = RARPasswordWin.GetPassword(item.RARFile)
		  ElseIf item = Nil Then
		    Dim item1 As RARItem = Me.RowTag(0)
		    If item1 <> Nil And item1.IsEncrypted Then pw = RARPasswordWin.GetPassword(item1.RARFile)
		  End If
		  
		  Select Case Left(hitItem.Text, 4)
		  Case "Extr"
		    Dim f As FolderItem
		    If item <> Nil Then
		      f = GetSaveFolderItem("", item.FileName)
		      index = item.Index
		    Else
		      f = SelectFolder()
		      index = -1
		    End If
		    If f <> Nil Then
		      If Archive.ExtractItem(Index, f, pw) Then
		        msg = "Extracted"
		      Else
		        msg = "RAR error " + Str(Archive.LastError) + ": " + UnRAR.FormatError(Archive.LastError)
		      End If
		    End If
		    
		  Case "Test"
		    If item <> Nil Then
		      index = item.Index
		    Else
		      index = -1
		    End If
		    If Archive.TestItem(Index, pw) Then
		      msg = "Test OK"
		    Else
		      msg = "RAR error " + Str(Archive.LastError) + ": " + UnRAR.FormatError(Archive.LastError)
		    End If
		    
		  End Select
		  If msg.Trim <> "" Then 
		    MsgBox(msg)
		    Return True
		  End If
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
		    Select Case UnRAR.PackingMethods(item.PackingMethod)
		    Case UnRAR.PackingMethods.Best
		      Listbox2.AddRow("Packing Method", "Best")
		    Case UnRAR.PackingMethods.Fast
		      Listbox2.AddRow("Packing Method", "Fast")
		    Case UnRAR.PackingMethods.Fastest
		      Listbox2.AddRow("Packing Method", "Fastest")
		    Case UnRAR.PackingMethods.Good
		      Listbox2.AddRow("Packing Method", "Good")
		    Case UnRAR.PackingMethods.Normal
		      Listbox2.AddRow("Packing Method", "Normal")
		    Case UnRAR.PackingMethods.Store
		      Listbox2.AddRow("Packing Method", "Store only")
		    Else
		      Listbox2.AddRow("Packing Method", Str(item.PackingMethod))
		    End Select
		    
		  End If
		End Sub
	#tag EndEvent
	#tag Event
		Function CompareRows(row1 as Integer, row2 as Integer, column as Integer, ByRef result as Integer) As Boolean
		  If column = 1 Then
		    Dim sz As Integer = Listbox1.CellTag(row1, 1)
		    Dim sz2 As Integer = Listbox1.CellTag(row2, 1)
		    If sz < sz2 Then
		      result = -1
		    ElseIf sz > sz2 Then
		      result = 1
		    Else
		      result = 0
		    End If
		    Return True
		  ElseIf column = 2 Then
		    Dim sz As Double = Listbox1.CellTag(row1, 2)
		    Dim sz2 As Double = Listbox1.CellTag(row2, 2)
		    If sz < sz2 Then
		      result = -1
		    ElseIf sz > sz2 Then
		      result = 1
		    Else
		      result = 0
		    End If
		    Return True
		  End If
		End Function
	#tag EndEvent
#tag EndEvents
#tag Events PushButton1
	#tag Event
		Sub Action()
		  If UnRAR.IsAvailable Then
		    Dim rar As FolderItem = GetOpenFolderItem(FileTypes1.All)
		    If rar <> Nil And rar.IsRARArchive Then
		      Listbox1.DeleteAllRows
		      Label1.Text = rar.AbsolutePath
		      Archive = New RARchive(rar)
		      AddHandler Archive.ProcessItem, WeakAddressOf Self.ProcessHandler
		      Dim count As Integer
		      Dim list() As RARItem = Archive.ListItems
		      count = UBound(list)
		      ProgressBar1.Maximum = count
		      ProgressBar1.Value = 0
		      For i As Integer = 0 To count
		        Dim item As RARItem = list(i)
		        If item.FileName.Trim = "" Then Break
		        If item.Directory Then
		          Listbox1.AddFolder(item.FileName)
		        Else
		          Dim d, sz, p As String
		          d = item.FileTime.SQLDateTime
		          sz = Format(item.UnpackedSize, "###,###,###,###")
		          p = Format(item.PackedSize * 100 / item.UnpackedSize, "#0.0#\%")
		          Listbox1.AddRow(item.FileName, sz, p, d)
		          Listbox1.CellTag(Listbox1.LastIndex, 1) = item.UnpackedSize
		          Listbox1.CellTag(Listbox1.LastIndex, 2) = item.PackedSize * 100 / item.UnpackedSize
		        End If
		        Listbox1.RowTag(Listbox1.LastIndex) = item
		      Next
		      Self.Title = "UnRar demo - " + rar.Name + "(" + Format(count, "###,###,##0") + " items)"
		    End If
		  Else
		    MsgBox(UnRAR.FormatError(UnRAR.ErrorRARUnavailable))
		  End If
		  
		Exception
		  If Archive <> Nil And Archive.LastError <> 0 Then
		    MsgBox("RAR error: " + Str(Archive.LastError))
		  End If
		End Sub
	#tag EndEvent
#tag EndEvents
