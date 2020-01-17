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
   MinHeight       =   585
   MinimizeButton  =   True
   MinWidth        =   600
   Placement       =   2
   Resizeable      =   True
   Title           =   "UnRar demo"
   Visible         =   True
   Width           =   6.0e+2
   Begin Listbox ArchList
      AutoDeactivate  =   True
      AutoHideScrollbars=   True
      Bold            =   ""
      Border          =   True
      ColumnCount     =   3
      ColumnsResizable=   True
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
      TabIndex        =   5
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
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   False
      Scope           =   0
      TabIndex        =   0
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
      TabIndex        =   7
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
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   False
      Maximum         =   100
      Scope           =   0
      TabPanelIndex   =   0
      Top             =   303
      Value           =   0
      Visible         =   True
      Width           =   600
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
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   False
      LockRight       =   True
      LockTop         =   False
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
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   False
      LockRight       =   True
      LockTop         =   False
      Scope           =   0
      TabIndex        =   2
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
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   False
      LockRight       =   True
      LockTop         =   False
      Scope           =   0
      TabIndex        =   4
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
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   False
      LockRight       =   True
      LockTop         =   False
      Scope           =   0
      TabIndex        =   3
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
   Begin TabPanel TabPanel1
      AutoDeactivate  =   True
      Bold            =   ""
      Enabled         =   True
      Height          =   267
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   ""
      Left            =   0
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   False
      Panels          =   ""
      Scope           =   0
      SmallTabs       =   ""
      TabDefinition   =   "Selected Item Info\rArchive Info"
      TabIndex        =   9
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   0
      TextUnit        =   0
      Top             =   318
      Underline       =   ""
      Value           =   1
      Visible         =   True
      Width           =   600
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
         Height          =   229
         HelpTag         =   ""
         Hierarchical    =   ""
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         InitialValue    =   "Property	Value"
         Italic          =   ""
         Left            =   6
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
         TabIndex        =   0
         TabPanelIndex   =   1
         TabStop         =   True
         TextFont        =   "System"
         TextSize        =   0
         TextUnit        =   0
         Top             =   349
         Underline       =   ""
         UseFocusRing    =   True
         Visible         =   True
         Width           =   588
         _ScrollWidth    =   -1
      End
      Begin Listbox ArchiveDetail
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
         Height          =   89
         HelpTag         =   ""
         Hierarchical    =   ""
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         InitialValue    =   "Property	Value"
         Italic          =   ""
         Left            =   5
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
         TabIndex        =   0
         TabPanelIndex   =   2
         TabStop         =   True
         TextFont        =   "System"
         TextSize        =   0
         TextUnit        =   0
         Top             =   350
         Underline       =   ""
         UseFocusRing    =   True
         Visible         =   True
         Width           =   588
         _ScrollWidth    =   -1
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
         InitialParent   =   "TabPanel1"
         Italic          =   ""
         Left            =   6
         LockBottom      =   ""
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   ""
         LockTop         =   True
         Multiline       =   ""
         Scope           =   0
         Selectable      =   False
         TabIndex        =   1
         TabPanelIndex   =   2
         Text            =   "Archive comment:"
         TextAlign       =   0
         TextColor       =   &h000000
         TextFont        =   "System"
         TextSize        =   0
         TextUnit        =   0
         Top             =   451
         Transparent     =   True
         Underline       =   ""
         Visible         =   True
         Width           =   100
      End
      Begin TextArea ArchComment
         AcceptTabs      =   ""
         Alignment       =   0
         AutoDeactivate  =   True
         AutomaticallyCheckSpelling=   True
         BackColor       =   &hFFFFFF
         Bold            =   ""
         Border          =   True
         DataField       =   ""
         DataSource      =   ""
         Enabled         =   True
         Format          =   ""
         Height          =   91
         HelpTag         =   ""
         HideSelection   =   True
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         Italic          =   ""
         Left            =   10
         LimitText       =   0
         LockBottom      =   ""
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   ""
         LockTop         =   True
         Mask            =   ""
         Multiline       =   True
         ReadOnly        =   True
         Scope           =   0
         ScrollbarHorizontal=   ""
         ScrollbarVertical=   True
         Styled          =   True
         TabIndex        =   2
         TabPanelIndex   =   2
         TabStop         =   True
         Text            =   ""
         TextColor       =   &h000000
         TextFont        =   "System"
         TextSize        =   0
         TextUnit        =   0
         Top             =   474
         Underline       =   ""
         UseFocusRing    =   True
         Visible         =   True
         Width           =   574
      End
   End
   Begin Timer CompleteTimer
      Height          =   32
      Index           =   -2147483648
      Left            =   612
      LockedInPosition=   False
      Mode            =   0
      Period          =   1
      Scope           =   0
      TabPanelIndex   =   0
      Top             =   -10
      Width           =   32
   End
   Begin Timer ProgressTimer
      Height          =   32
      Index           =   -2147483648
      Left            =   612
      LockedInPosition=   False
      Mode            =   0
      Period          =   200
      Scope           =   0
      TabPanelIndex   =   0
      Top             =   23
      Width           =   32
   End
   Begin PushButton AbortBtn
      AutoDeactivate  =   True
      Bold            =   ""
      ButtonStyle     =   0
      Cancel          =   ""
      Caption         =   "Abort"
      Default         =   ""
      Enabled         =   True
      Height          =   22
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   ""
      Left            =   180
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   False
      LockRight       =   True
      LockTop         =   False
      Scope           =   0
      TabIndex        =   10
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   0
      TextUnit        =   0
      Top             =   274
      Underline       =   ""
      Visible         =   False
      Width           =   80
   End
   Begin Timer GetPasswordTimer
      Height          =   32
      Index           =   -2147483648
      Left            =   880
      LockedInPosition=   False
      Mode            =   0
      Period          =   1
      Scope           =   0
      TabPanelIndex   =   0
      Top             =   29
      Width           =   32
   End
   Begin Timer ChangeVolumeTimer
      Height          =   32
      Index           =   -2147483648
      Left            =   880
      LockedInPosition=   False
      Mode            =   0
      Period          =   1
      Scope           =   0
      TabPanelIndex   =   0
      Top             =   67
      Width           =   32
   End
End
#tag EndWindow

#tag WindowCode
	#tag Method, Flags = &h21
		Private Function ChangeVolumeHandler(Sender As UnRAR.IteratorEx, VolumeNumber As Integer, ByRef NextVolume As FolderItem) As Boolean
		  If App.CurrentThread = Nil Then
		    mNextVolume = RARVolumeSelect.GetVolume(Sender.ArchiveFile.Name, VolumeNumber, NextVolume)
		  Else
		    mVolumeNumber = VolumeNumber
		    mNextVolume = NextVolume
		    mWaiting = True
		    ChangeVolumeTimer.Mode = Timer.ModeSingle
		    Do Until Not mWaiting
		      App.YieldToNextThread
		    Loop
		  End If
		  
		  If mNextVolume <> Nil Then
		    NextVolume = mNextVolume
		    Return True
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function GetPasswordHandler(Sender As UnRAR.IteratorEx, ByRef ArchivePassword As String) As Boolean
		  If App.CurrentThread = Nil Then
		    mPassword = RARPasswordWin.GetPassword(Sender.ArchiveFile)
		  Else
		    mWaiting = True
		    GetPasswordTimer.Mode = Timer.ModeSingle
		    Do Until Not mWaiting
		      App.YieldToNextThread
		    Loop
		  End If
		  ArchivePassword = mPassword
		  Return True
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function ProcessDataHandler(Sender As UnRAR.IteratorEx, NewData As MemoryBlock, Length As Integer) As Boolean
		  #pragma Unused Sender
		  #pragma Unused NewData
		  #pragma Unused Length
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub RunExtractAll(Sender As Thread)
		  #pragma Unused Sender
		  Do
		    mCurrentIndex = mCurrentIndex + 1
		    If Not Archive.MoveNext(UnRAR.RAR_EXTRACT, mOutputDir) Then Exit Do
		    App.YieldToNextThread
		  Loop Until Archive.LastError <> 0
		  CompleteTimer.Mode = Timer.ModeSingle
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub RunExtractSingle(Sender As Thread)
		  #pragma Unused Sender
		  
		  Do Until Archive.CurrentItem.Index = mSelectedIndex
		    mCurrentIndex = mCurrentIndex + 1
		    App.YieldToNextThread
		  Loop Until Not Archive.MoveNext(UnRAR.RAR_SKIP)
		  Call Archive.MoveNext(UnRAR.RAR_Extract, mOutputDir)
		  CompleteTimer.Mode = Timer.ModeSingle
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub RunTestAll(Sender As Thread)
		  #pragma Unused Sender
		  Do
		    mCurrentIndex = mCurrentIndex + 1
		    If Not Archive.MoveNext(UnRAR.RAR_TEST) Then Exit Do
		  Loop
		  CompleteTimer.Mode = Timer.ModeSingle
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub RunTestSingle(Sender As Thread)
		  #pragma Unused Sender
		  
		  Do Until Archive.CurrentItem.Index = mSelectedIndex
		    mCurrentIndex = mCurrentIndex + 1
		  Loop Until Not Archive.MoveNext(UnRAR.RAR_SKIP)
		  Call Archive.MoveNext(UnRAR.RAR_TEST)
		  CompleteTimer.Mode = Timer.ModeSingle
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function VolumeChangedHandler(Sender As UnRAR.IteratorEx, VolumeNumber As Integer, NextVolume As FolderItem) As Boolean
		  #pragma Unused Sender
		  #If DebugBuild Then
		    System.DebugLog("Changed volume to #" + Str(VolumeNumber) + "(" + NextVolume.AbsolutePath_ + ")")
		  #endif
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
		Private mCurrentIndex As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mNextVolume As FolderItem
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mOutputDir As FolderItem
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mPassword As String
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mSelectedIndex As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mVolumeNumber As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mWaiting As Boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mWorker As Thread
	#tag EndProperty

	#tag Property, Flags = &h21
		Private SizeAll As UInt64
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected SizeCompressed As UInt64
	#tag EndProperty

	#tag Property, Flags = &h21
		Private SizeNow As UInt64
	#tag EndProperty


#tag EndWindowCode

#tag Events ArchList
	#tag Event
		Sub Change()
		  ItemDetail.DeleteAllRows
		  mCurrentIndex = Me.ListIndex
		  If mCurrentIndex > -1 Then
		    Dim item As UnRAR.ArchiveEntry = Me.RowTag(Me.ListIndex)
		    If item = Nil Then Return
		    ItemDetail.AddRow("File Name",  item.FileName)
		    ItemDetail.AddRow("Time",  item.FileTime.SQLDateTime)
		    ItemDetail.AddRow("Archive", item.RARFile.AbsolutePath_)
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
		  If Not UnRAR.IsAvailable Then
		    MsgBox(UnRAR.FormatError(UnRAR.ErrorRARUnavailable))
		    Return
		  End If
		  Dim rar As FolderItem = GetOpenFolderItem(FileTypes1.All)
		  If rar = Nil Or Not rar.IsRARArchive Then
		    MsgBox("Not a RAR archive.")
		    Return
		  End If
		  
		  ArchList.DeleteAllRows
		  ArchivePath.Text = rar.AbsolutePath_
		  If Archive <> Nil Then Archive.Close
		  Archive = New UnRAR.IteratorEx(rar)
		  AddHandler Archive.GetPassword, WeakAddressOf GetPasswordHandler
		  AddHandler Archive.ChangeVolume, WeakAddressOf ChangeVolumeHandler
		  AddHandler Archive.ProcessData, WeakAddressOf ProcessDataHandler
		  AddHandler Archive.VolumeChanged, WeakAddressOf VolumeChangedHandler
		  Count = 0
		  SizeAll = 0
		  SizeNow = 0
		  SizeCompressed = 0
		  
		  Do
		    Dim d, sz, p As String
		    d = Archive.CurrentItem.FileTime.SQLDateTime
		    If Not Archive.CurrentItem.Directory Then
		      SizeAll = SizeAll + Archive.CurrentItem.UnpackedSize
		      SizeCompressed = SizeCompressed + Archive.CurrentItem.PackedSize
		      sz = Format(Archive.CurrentItem.UnpackedSize, "###,###,###,###")
		      p = Format(Archive.CurrentItem.PackedSize * 100 /Archive.CurrentItem.UnpackedSize, "#0.0#\%")
		    End If
		    ArchList.AddRow(Archive.CurrentItem.FileName, sz, p, d)
		    ArchList.CellTag(ArchList.LastIndex, 1) = Archive.CurrentItem.UnpackedSize
		    ArchList.CellTag(ArchList.LastIndex, 2) = Archive.CurrentItem.PackedSize * 100 / Archive.CurrentItem.UnpackedSize
		    ArchList.RowTag(ArchList.LastIndex) = Archive.CurrentItem
		    count = count + 1
		  Loop Until Not Archive.MoveNext(UnRAR.RAR_SKIP)
		  ArchComment.Text = Archive.ArchiveComment
		  ArchiveDetail.DeleteAllRows
		  Dim a As UnRAR.ArchiveEntry
		  Try
		    a = ArchList.RowTag(0)
		    ArchiveDetail.AddRow("Archive", Archive.ArchiveFile.AbsolutePath_)
		    ArchiveDetail.AddRow("Total files", Format(count, "###,###,##0"))
		    ArchiveDetail.AddRow("Total size", Format(SizeAll, "###,###,###,###,##0"))
		    ArchiveDetail.AddRow("Packed size", Format(SizeCompressed, "###,###,###,###,##0"))
		    ArchiveDetail.AddRow("Compression ratio", Format(SizeCompressed * 100 / SizeAll, "##0.0#") + "%")
		    ArchiveDetail.AddRow("Solid", Str(a.IsSolid))
		    ArchiveDetail.AddRow("Encrypted", Str(a.IsEncrypted))
		    ArchiveDetail.AddRow("Host OS", a.HostOS)
		  Catch
		    //meh
		  End Try
		  Archive.Reset()
		  
		  
		  Self.Title = "UnRar demo - " + rar.Name + "(" + Format(count, "###,###,##0") + " items)"
		  ExtractAll.Enabled = True
		  ExtractItem.Enabled = False
		  TestAll.Enabled = True
		  TestItem.Enabled = False
		  ProgressBar1.Maximum = Count
		  
		Exception Err As RuntimeException
		  If Archive <> Nil And Archive.LastError <> 0 Then
		    MsgBox("RAR error " + Str(Archive.LastError) + ": " + UnRAR.FormatError(Archive.LastError))
		  Else
		    MsgBox("RAR error " + Str(Err.ErrorNumber) + ": " + UnRAR.FormatError(Err.ErrorNumber))
		  End If
		  
		  
		  
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events TestItem
	#tag Event
		Sub Action()
		  ProgressBar1.Value = 0
		  mCurrentIndex = 0
		  Dim item As UnRAR.ArchiveEntry = ArchList.RowTag(ArchList.ListIndex)
		  If item = Nil Then Return
		  
		  mWorker = New Thread
		  AddHandler mWorker.Run, WeakAddressOf RunTestSingle
		  mWorker.Run()
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events TestAll
	#tag Event
		Sub Action()
		  ProgressBar1.Value = 0
		  mCurrentIndex = 0
		  ProgressTimer.Mode = Timer.ModeMultiple
		  
		  mWorker = New Thread
		  AddHandler mWorker.Run, WeakAddressOf RunTestAll
		  mWorker.Run()
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ExtractAll
	#tag Event
		Sub Action()
		  mOutputDir = SelectFolder()
		  If mOutputDir = Nil Then Return
		  
		  ProgressBar1.Value = 0
		  mCurrentIndex = 0
		  ProgressTimer.Mode = Timer.ModeMultiple
		  
		  mWorker = New Thread
		  mWorker.Priority = Thread.HighPriority
		  AddHandler mWorker.Run, WeakAddressOf RunExtractAll
		  mWorker.Run()
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ExtractItem
	#tag Event
		Sub Action()
		  Dim item As UnRAR.ArchiveEntry = ArchList.RowTag(ArchList.ListIndex)
		  If item = Nil Then Return
		  mOutputDir = GetSaveFolderItem("", item.FileName)
		  If mOutputDir = Nil Then Return
		  
		  ProgressBar1.Value = 0
		  mCurrentIndex = 0
		  ProgressTimer.Mode = Timer.ModeMultiple
		  
		  mWorker = New Thread
		  AddHandler mWorker.Run, WeakAddressOf RunExtractSingle
		  mWorker.Run()
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events CompleteTimer
	#tag Event
		Sub Action()
		  ProgressBar1.Value = 0
		  ProgressTimer.Mode = Timer.ModeOff
		  AbortBtn.Visible = False
		  mWorker = Nil
		  If Archive.LastError = UnRAR.ErrorEndArchive Or Archive.LastError = 0 Then
		    MsgBox("Operation succeeded.")
		  Else
		    MsgBox("RAR error " + Str(Archive.LastError) + ": " + UnRAR.FormatError(Archive.LastError))
		  End If
		  Archive.Reset()
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ProgressTimer
	#tag Event
		Sub Action()
		  ProgressBar1.Value = mCurrentIndex * 100 / Count
		  AbortBtn.Visible = True
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events AbortBtn
	#tag Event
		Sub Action()
		  If mWorker <> Nil Then
		    mWorker.Kill
		    CompleteTimer.Mode = Timer.ModeSingle
		  End If
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events GetPasswordTimer
	#tag Event
		Sub Action()
		  mPassword = RARPasswordWin.GetPassword(Archive.ArchiveFile)
		  mWaiting = False
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ChangeVolumeTimer
	#tag Event
		Sub Action()
		  mNextVolume = RARVolumeSelect.GetVolume(Archive.ArchiveFile.Name, mVolumeNumber, mNextVolume)
		  mWaiting = False
		End Sub
	#tag EndEvent
#tag EndEvents
