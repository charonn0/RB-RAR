#tag Window
Begin Window Window1
   BackColor       =   16777215
   Backdrop        =   ""
   CloseButton     =   True
   Composite       =   False
   Frame           =   0
   FullScreen      =   False
   HasBackColor    =   False
   Height          =   400
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
   Width           =   600
End
#tag EndWindow

#tag WindowCode
	#tag Event
		Sub Open()
		  Dim r As New UnRar.RARchive(GetOpenFolderItem(""))
		  'Call r.ExtractAll(SpecialFolder.Desktop.Child("Output"))
		  'Dim g As FolderItem = r.ExtractItem(3)
		  'g.MoveFileTo(SpecialFolder.Desktop)
		  'If r.TestAll Then
		  'Break
		  'Else
		  'Break
		  'End If
		  Dim count As Integer = r.Count
		  For i As Integer = 1 To count
		    Dim ri As UnRAR.RARItem = r.Item(i)
		    Dim a, c, n, os As String
		    Dim ver As Double
		    Dim cry, sol As Boolean
		    Dim CR, att, fl, pkd, upkd, m As Integer
		    a = ri.Archive.AbsolutePath
		    c = ri.Comment
		    cr = ri.CRC32
		    cry = ri.IsEncrypted
		    att = ri.FileAttributes
		    n = ri.FileName
		    fl = ri.Flags
		    os = ri.HostOS
		    sol = ri.IsSolid
		    ver = ri.MinimumVersion
		    pkd = ri.PackedSize
		    m = ri.PackingMethod
		    upkd = ri.UnpackedSize
		    Break
		  Next
		End Sub
	#tag EndEvent


#tag EndWindowCode

