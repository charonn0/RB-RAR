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
   Begin Listbox ArchList
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
      Height          =   242
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
   Begin PushButton OpenRAR
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
      Left            =   6
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
      Top             =   274
      Underline       =   ""
      Visible         =   True
      Width           =   80
   End
   Begin GroupBox GroupBox1
      AutoDeactivate  =   True
      Bold            =   ""
      Caption         =   "Selected Item Properties"
      Enabled         =   True
      Height          =   263
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
      Top             =   313
      Underline       =   ""
      Visible         =   True
      Width           =   588
   End
   Begin Label ArchivePath
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
      Width           =   588
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
      Top             =   303
      Value           =   0
      Visible         =   True
      Width           =   600
   End
   Begin Listbox ItemDetail
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
      Height          =   230
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
      Top             =   335
      Underline       =   ""
      UseFocusRing    =   True
      Visible         =   True
      Width           =   560
      _ScrollWidth    =   -1
   End
   Begin PushButton TestItem
      AutoDeactivate  =   True
      Bold            =   ""
      ButtonStyle     =   0
      Cancel          =   ""
      Caption         =   "Test Item"
      Default         =   ""
      Enabled         =   False
      Height          =   22
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   ""
      Left            =   262
      LockBottom      =   ""
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   ""
      LockTop         =   True
      Scope           =   0
      TabIndex        =   5
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   0
      TextUnit        =   0
      Top             =   274
      Underline       =   ""
      Visible         =   True
      Width           =   80
   End
   Begin PushButton TestAll
      AutoDeactivate  =   True
      Bold            =   ""
      ButtonStyle     =   0
      Cancel          =   ""
      Caption         =   "Test All"
      Default         =   ""
      Enabled         =   False
      Height          =   22
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   ""
      Left            =   346
      LockBottom      =   ""
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   ""
      LockTop         =   True
      Scope           =   0
      TabIndex        =   6
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   0
      TextUnit        =   0
      Top             =   274
      Underline       =   ""
      Visible         =   True
      Width           =   80
   End
   Begin PushButton ExtractAll
      AutoDeactivate  =   True
      Bold            =   ""
      ButtonStyle     =   0
      Cancel          =   ""
      Caption         =   "Extract All"
      Default         =   ""
      Enabled         =   False
      Height          =   22
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   ""
      Left            =   514
      LockBottom      =   ""
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   ""
      LockTop         =   True
      Scope           =   0
      TabIndex        =   7
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   0
      TextUnit        =   0
      Top             =   274
      Underline       =   ""
      Visible         =   True
      Width           =   80
   End
   Begin PushButton ExtractItem
      AutoDeactivate  =   True
      Bold            =   ""
      ButtonStyle     =   0
      Cancel          =   ""
      Caption         =   "Extract Item"
      Default         =   ""
      Enabled         =   False
      Height          =   22
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   ""
      Left            =   430
      LockBottom      =   ""
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   ""
      LockTop         =   True
      Scope           =   0
      TabIndex        =   8
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   0
      TextUnit        =   0
      Top             =   274
      Underline       =   ""
      Visible         =   True
      Width           =   80
   End
End
#tag EndWindow

#tag WindowCode
	#tag Method, Flags = &h0
		Function ProgressHandler(Sender As RARchive, NextItem As RARItem) As Boolean
		  #pragma Unused Sender
		  ProgressBar1.Value = NextItem.Index + 1
		End Function
	#tag EndMethod


	#tag Property, Flags = &h1
		Protected Archive As UnRAR.RARchive
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected LastTop As Integer
	#tag EndProperty


#tag EndWindowCode

#tag Events ArchList
	#tag Event
		Sub Change()
		  ItemDetail.DeleteAllRows
		  If Me.ListIndex > -1 Then
		    Dim item As RARItem = Me.RowTag(Me.ListIndex)
		    If item = Nil Then Return
		    ItemDetail.AddRow("File Name",  item.FileName)
		    ItemDetail.AddRow("Time",  item.FileTime.SQLDateTime)
		    ItemDetail.AddRow("Archive", item.RARFile.AbsolutePath)
		    ItemDetail.AddRow("Volume", item.VolumeName)
		    ItemDetail.AddRow("Comment",  item.Comment)
		    ItemDetail.AddRow("Packed size",  Format(item.PackedSize, "###,###,###,###"))
		    ItemDetail.AddRow("Unpacked size",  Format(item.UnpackedSize, "###,###,###,###"))
		    ItemDetail.AddRow("Index",  Str(item.Index))
		    ItemDetail.AddRow("Encrypted",  Str(item.IsEncrypted))
		    ItemDetail.AddRow("Solid",  Str(item.IsSolid))
		    ItemDetail.AddRow("CRC32",  "0x" + Left(Hex(item.CRC32) + "00000000", 8))
		    ItemDetail.AddRow("Directory",  Str(item.Directory))
		    ItemDetail.AddRow("Attributes",  "0x" + Left(Hex(item.FileAttributes) + "00000000", 8))
		    ItemDetail.AddRow("Flags",  "0x" + Left(Hex(item.Flags) + "00000000", 8))
		    ItemDetail.AddRow("OS",  item.HostOS)
		    ItemDetail.AddRow("Min. Version",  Format(item.MinimumVersion, "#0.0#"))
		    Select Case UnRAR.PackingMethods(item.PackingMethod)
		    Case UnRAR.PackingMethods.Best
		      ItemDetail.AddRow("Packing Method", "Best")
		    Case UnRAR.PackingMethods.Fast
		      ItemDetail.AddRow("Packing Method", "Fast")
		    Case UnRAR.PackingMethods.Fastest
		      ItemDetail.AddRow("Packing Method", "Fastest")
		    Case UnRAR.PackingMethods.Good
		      ItemDetail.AddRow("Packing Method", "Good")
		    Case UnRAR.PackingMethods.Normal
		      ItemDetail.AddRow("Packing Method", "Normal")
		    Case UnRAR.PackingMethods.Store
		      ItemDetail.AddRow("Packing Method", "Store only")
		    Else
		      ItemDetail.AddRow("Packing Method", Str(item.PackingMethod))
		    End Select
		    
		    ExtractAll.Enabled = True
		    ExtractItem.Enabled = True
		    TestAll.Enabled = True
		    TestItem.Enabled = True
		  Else
		    ExtractAll.Enabled = Archive <> Nil
		    ExtractItem.Enabled = False
		    TestAll.Enabled = Archive <> Nil
		    TestItem.Enabled = False
		  End If
		End Sub
	#tag EndEvent
	#tag Event
		Function CompareRows(row1 as Integer, row2 as Integer, column as Integer, ByRef result as Integer) As Boolean
		  If column = 1 Then
		    Dim sz As Integer = ArchList.CellTag(row1, 1)
		    Dim sz2 As Integer = ArchList.CellTag(row2, 1)
		    If sz < sz2 Then
		      result = -1
		    ElseIf sz > sz2 Then
		      result = 1
		    Else
		      result = 0
		    End If
		    Return True
		  ElseIf column = 2 Then
		    Dim sz As Double = ArchList.CellTag(row1, 2)
		    Dim sz2 As Double = ArchList.CellTag(row2, 2)
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
#tag Events OpenRAR
	#tag Event
		Sub Action()
		  If UnRAR.IsAvailable Then
		    Dim rar As FolderItem = GetOpenFolderItem(FileTypes1.All)
		    If rar <> Nil And rar.IsRARArchive Then
		      ArchList.DeleteAllRows
		      ArchivePath.Text = rar.AbsolutePath
		      Archive = New RARchive(rar)
		      AddHandler Archive.OperationProgress, WeakAddressOf Self.ProgressHandler
		      Dim count As Integer
		      Dim list() As RARItem = Archive.ListItems
		      count = UBound(list)
		      ProgressBar1.Maximum = count
		      ProgressBar1.Value = 0
		      For i As Integer = 0 To count
		        Dim item As RARItem = list(i)
		        If item.FileName.Trim = "" Then Break
		        If item.Directory Then
		          ArchList.AddFolder(item.FileName)
		        Else
		          Dim d, sz, p As String
		          d = item.FileTime.SQLDateTime
		          sz = Format(item.UnpackedSize, "###,###,###,###")
		          p = Format(item.PackedSize * 100 / item.UnpackedSize, "#0.0#\%")
		          ArchList.AddRow(item.FileName, sz, p, d)
		          ArchList.CellTag(ArchList.LastIndex, 1) = item.UnpackedSize
		          ArchList.CellTag(ArchList.LastIndex, 2) = item.PackedSize * 100 / item.UnpackedSize
		        End If
		        ArchList.RowTag(ArchList.LastIndex) = item
		      Next
		      Self.Title = "UnRar demo - " + rar.Name + "(" + Format(count + 1, "###,###,##0") + " items)"
		      ExtractAll.Enabled = True
		      ExtractItem.Enabled = False
		      TestAll.Enabled = True
		      TestItem.Enabled = False
		      ProgressBar1.Maximum = ArchList.ListCount - 1
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
#tag Events TestItem
	#tag Event
		Sub Action()
		  Dim item As RARItem = ArchList.RowTag(ArchList.ListIndex)
		  Dim pw As String
		  
		  If item <> Nil Then
		    If item.IsEncrypted Then pw = RARPasswordWin.GetPassword(item.RARFile)
		    If Archive.TestItem(item.Index, pw) Then
		      MsgBox("Test OK")
		    Else
		      MsgBox("RAR error " + Str(Archive.LastError) + ": " + UnRAR.FormatError(Archive.LastError))
		    End If
		  End If
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events TestAll
	#tag Event
		Sub Action()
		  Dim item As RARItem = ArchList.RowTag(0)
		  Dim pw As String
		  
		  If item <> Nil Then
		    If item.IsEncrypted Then pw = RARPasswordWin.GetPassword(item.RARFile)
		    If Archive.TestItem(-1, pw) Then
		      MsgBox("Test OK")
		    Else
		      MsgBox("RAR error " + Str(Archive.LastError) + ": " + UnRAR.FormatError(Archive.LastError))
		    End If
		  End If
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ExtractAll
	#tag Event
		Sub Action()
		  Dim item As RARItem = ArchList.RowTag(0)
		  Dim pw As String
		  Dim f As FolderItem
		  
		  If item <> Nil Then
		    If item.IsEncrypted Then pw = RARPasswordWin.GetPassword(item.RARFile)
		    f = SelectFolder()
		    If f <> Nil And Archive.ExtractItem(-1, f, pw) Then
		      MsgBox("Extract OK")
		    ElseIf f <> Nil Then
		      MsgBox("RAR error " + Str(Archive.LastError) + ": " + UnRAR.FormatError(Archive.LastError))
		    End If
		  End If
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ExtractItem
	#tag Event
		Sub Action()
		  Dim item As RARItem = ArchList.RowTag(ArchList.ListIndex)
		  Dim pw As String
		  Dim f As FolderItem
		  
		  If item <> Nil Then
		    If item.IsEncrypted Then pw = RARPasswordWin.GetPassword(item.RARFile)
		    f = GetSaveFolderItem("", item.FileName)
		    If f <> Nil And Archive.ExtractItem(item.Index, f, pw) Then
		      MsgBox("Extract OK")
		    ElseIf f <> Nil Then
		      MsgBox("RAR error " + Str(Archive.LastError) + ": " + UnRAR.FormatError(Archive.LastError))
		    End If
		  End If
		End Sub
	#tag EndEvent
#tag EndEvents
