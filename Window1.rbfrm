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
		    Call ri.FileTime
		    Break
		  Next
		End Sub
	#tag EndEvent


#tag EndWindowCode

