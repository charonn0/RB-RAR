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
	#tag Method, Flags = &h21
		Private Function ChangeVolumeHandler(Sender As UnRAR.IteratorEx, VolumeNumber As Integer, ByRef NextVolume As FolderItem) As Boolean
		  Dim f As FolderItem = RARVolumeSelect.GetVolume(Sender.ArchiveFile.Name, VolumeNumber, NextVolume)
		  If f <> Nil Then
		    NextVolume = f
		    Return True
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function GetPasswordHandler(Sender As UnRAR.IteratorEx, ByRef ArchivePassword As String) As Boolean
		  ArchivePassword = RARPasswordWin.GetPassword(Sender.ArchiveFile)
		  Return True
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function ProcessDataHandler(Sender As UnRAR.IteratorEx, NewData As MemoryBlock, Length As Integer) As Boolean
		  #pragma Unused Sender
		  #pragma Unused NewData
		  
		  SizeNow = SizeNow + Length
		  ProgressBar1.Value = SizeNow * 100 / SizeAll
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function VolumeChangedHandler(Sender As UnRAR.IteratorEx, VolumeNumber As Integer, NextVolume As FolderItem) As Boolean
		  System.DebugLog("Changed volume to #" + Str(VolumeNumber) + "(" + NextVolume.AbsolutePath + ")")
		End Function
	#tag EndMethod


	#tag Property, Flags = &h1
		Protected Archive As UnRAR.IteratorEx
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected Count As Integer
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected LastTop As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private SizeAll As UInt64
	#tag EndProperty

	#tag Property, Flags = &h21
		Private SizeNow As UInt64
	#tag EndProperty


#tag EndWindowCode

#tag Events ArchList
	#tag Event
		Sub Change()
		  ItemDetail.DeleteAllRows
		  If Me.ListIndex > -1 Then
		    Dim item As UnRAR.ArchiveEntry = Me.RowTag(Me.ListIndex)
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
		      If Archive <> Nil Then Archive.Close
		      Archive = New UnRAR.IteratorEx(rar)
		      AddHandler Archive.GetPassword, WeakAddressOf GetPasswordHandler
		      AddHandler Archive.ChangeVolume, WeakAddressOf ChangeVolumeHandler
		      AddHandler Archive.ProcessData, WeakAddressOf ProcessDataHandler
		      AddHandler Archive.VolumeChanged, WeakAddressOf VolumeChangedHandler
		      Count = 0
		      SizeAll = 0
		      SizeNow = 0
		      Do
		        Dim d, sz, p As String
		        d = Archive.CurrentItem.FileTime.SQLDateTime
		        If Not Archive.CurrentItem.Directory Then
		          SizeAll = SizeAll + Archive.CurrentItem.UnpackedSize
		          sz = Format(Archive.CurrentItem.UnpackedSize, "###,###,###,###")
		          p = Format(Archive.CurrentItem.PackedSize * 100 /Archive.CurrentItem.UnpackedSize, "#0.0#\%")
		        End If
		        ArchList.AddRow(Archive.CurrentItem.FileName, sz, p, d)
		        ArchList.CellTag(ArchList.LastIndex, 1) = Archive.CurrentItem.UnpackedSize
		        ArchList.CellTag(ArchList.LastIndex, 2) = Archive.CurrentItem.PackedSize * 100 / Archive.CurrentItem.UnpackedSize
		        ArchList.RowTag(ArchList.LastIndex) = Archive.CurrentItem
		        count = count + 1
		      Loop Until Not Archive.MoveNext(UnRAR.RAR_SKIP)
		      Archive.Reset()
		      Self.Title = "UnRar demo - " + rar.Name + "(" + Format(count + 1, "###,###,##0") + " items)"
		      ExtractAll.Enabled = True
		      ExtractItem.Enabled = False
		      TestAll.Enabled = True
		      TestItem.Enabled = False
		    End If
		    ProgressBar1.Maximum = Count
		  Else
		    MsgBox(UnRAR.FormatError(UnRAR.ErrorRARUnavailable))
		  End If
		  
		Exception Err As RuntimeException
		  If Archive <> Nil And Archive.LastError <> 0 Then
		    MsgBox("RAR error: " + Str(Archive.LastError) + " " + UnRAR.FormatError(Archive.LastError))
		  Else
		    MsgBox("RAR error: " + Str(Err.ErrorNumber) + " " + UnRAR.FormatError(Err.ErrorNumber))
		  End If
		  
		  
		  
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events TestItem
	#tag Event
		Sub Action()
		  ProgressBar1.Value = 0
		  SizeNow = 0
		  Dim item As UnRAR.ArchiveEntry = ArchList.RowTag(ArchList.ListIndex)
		  If item = Nil Then Return
		  
		  Do Until Archive.CurrentItem.Index = item.Index
		    App.YieldToNextThread
		  Loop Until Not Archive.MoveNext(UnRAR.RAR_SKIP)
		  
		  If Archive.MoveNext(UnRAR.RAR_TEST) Or Archive.LastError = UnRAR.ErrorEndArchive Then
		    MsgBox("Test OK")
		  Else
		    MsgBox("RAR error " + Str(Archive.LastError) + ": " + UnRAR.FormatError(Archive.LastError))
		  End If
		  Archive.Reset()
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events TestAll
	#tag Event
		Sub Action()
		  ProgressBar1.Value = 0
		  SizeNow = 0
		  Do
		    If Not Archive.MoveNext(UnRAR.RAR_TEST) Then Exit Do
		  Loop
		  If Archive.LastError = UnRAR.ErrorEndArchive Then
		    MsgBox("Test OK")
		  Else
		    MsgBox("RAR error " + Str(Archive.LastError) + ": " + UnRAR.FormatError(Archive.LastError))
		  End If
		  Archive.Reset()
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ExtractAll
	#tag Event
		Sub Action()
		  ProgressBar1.Value = 0
		  SizeNow = 0
		  Dim f As FolderItem = SelectFolder()
		  If f = Nil Then Return
		  Do
		    If Not Archive.MoveNext(UnRAR.RAR_EXTRACT, f) Then Exit Do
		  Loop Until Archive.LastError <> 0
		  If Archive.LastError = UnRAR.ErrorEndArchive Then
		    MsgBox("Extract OK")
		  Else
		    MsgBox("RAR error " + Str(Archive.LastError) + ": " + UnRAR.FormatError(Archive.LastError))
		  End If
		  Archive.Reset()
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ExtractItem
	#tag Event
		Sub Action()
		  ProgressBar1.Value = 0
		  SizeNow = 0
		  Dim item As UnRAR.ArchiveEntry = ArchList.RowTag(ArchList.ListIndex)
		  If item = Nil Then Return
		  Dim f As FolderItem = GetSaveFolderItem("", item.FileName)
		  If f = Nil Then Return
		  
		  Do Until Archive.CurrentItem.Index = item.Index
		    App.YieldToNextThread
		  Loop Until Not Archive.MoveNext(UnRAR.RAR_SKIP)
		  If Archive.MoveNext(UnRAR.RAR_Extract, f) Or Archive.LastError = UnRAR.ErrorEndArchive Then
		    MsgBox("Extract OK")
		  Else
		    MsgBox("RAR error " + Str(Archive.LastError) + ": " + UnRAR.FormatError(Archive.LastError))
		  End If
		  Archive.Reset()
		End Sub
	#tag EndEvent
#tag EndEvents
