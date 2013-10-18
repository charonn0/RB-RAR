#tag Class
Protected Class App
Inherits Application
	#tag Event
		Sub Open()
		  Dim f As FolderItem = GetOpenFolderItem("")
		  Arch = New RARchiveEx(f)
		  AddHandler Arch.ProcessData, WeakAddressOf App.RARProcess
		  AddHandler Arch.PasswordPrompt, WeakAddressOf App.RARPassword
		  AddHandler Arch.ChangeVolume, WeakAddressOf App.RARVolume
		  Arch.TestAll
		  
		  While Arch.LastError = 0
		    DoEvents
		  Wend
		  Quit
		End Sub
	#tag EndEvent


	#tag Method, Flags = &h21
		Private Function RARPassword(Sender As RARChiveEx) As String
		  Break
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function RARProcess(Sender As RARChiveEx, Data As Readable, Length As Integer, Header As RARItemEx) As Boolean
		  Break
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function RARVolume(Sender As RARChiveEx, ByRef NextVolume As FolderItem) As Boolean
		  Break
		End Function
	#tag EndMethod


	#tag Property, Flags = &h21
		Private Arch As RARchiveEx
	#tag EndProperty


	#tag Constant, Name = kEditClear, Type = String, Dynamic = False, Default = \"&Delete", Scope = Public
		#Tag Instance, Platform = Windows, Language = Default, Definition  = \"&Delete"
		#Tag Instance, Platform = Linux, Language = Default, Definition  = \"&Delete"
	#tag EndConstant

	#tag Constant, Name = kFileQuit, Type = String, Dynamic = False, Default = \"&Quit", Scope = Public
		#Tag Instance, Platform = Windows, Language = Default, Definition  = \"E&xit"
	#tag EndConstant

	#tag Constant, Name = kFileQuitShortcut, Type = String, Dynamic = False, Default = \"", Scope = Public
		#Tag Instance, Platform = Mac OS, Language = Default, Definition  = \"Cmd+Q"
		#Tag Instance, Platform = Linux, Language = Default, Definition  = \"Ctrl+Q"
	#tag EndConstant


	#tag ViewBehavior
	#tag EndViewBehavior
End Class
#tag EndClass
