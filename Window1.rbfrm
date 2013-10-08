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
		  Dim f As FolderItem = GetOpenFolderItem("")
		  Dim r As RARchive = RARchive.Open(f)
		  'Dim i As Integer = r.ItemCount
		  Dim ri As RARItem = r.Item(1)
		  Dim a, c, n, os As String
		  Dim ver As Double
		  Dim cry, sol As Boolean
		  Dim CR, att, fl, pkd, upkd, m As Integer
		  a = ri.ArchiveName
		  c = ri.Comment
		  cr = ri.CRC32
		  cry = ri.Encrypted
		  att = ri.FileAttributes
		  n = ri.FileName
		  fl = ri.Flags
		  os = ri.HostOS
		  sol = ri.IsSolid
		  ver = ri.MinimumVersion
		  pkd = ri.PackedSize
		  m = ri.PackingMethod
		  upkd = ri.UnpackedSize
		  Call r.ExtractItem(1, SpecialFolder.Desktop.Child(ri.FileName))
		  Break
		End Sub
	#tag EndEvent


#tag EndWindowCode

