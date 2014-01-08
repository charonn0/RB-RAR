This project implements bindings for RealStudio and Xojo Win32 applications to extract RAR archives 
using the unrar library available here: http://www.rarlab.com/rar/UnRARDLL.exe. The unrar.dll file 
must be stored on Windows' search path or in the same directory as your executable.

The `RARchive` class represents an entire WinRAR archive. Pass the RAR file to the class constructor:

    Dim rarfile As FolderItem ' assume a valid RAR file
    Dim myRARchive As New RARChive(rarfile)
	
Once a `RARchive` is instantiated, you can get metadata on any file in the archive by using the 
`RARchive.Item` method. This method accepts the index of the item in the archive and returns a `RARItem`
representing that item.

Interacting with individual `RARitems` in a `RARchive` is index-based. Continuing the above code sample, the
following code will extract the first item (index=0) from the archive to a user-selected location:
  
    Dim item As RARItem = myRARchive.Item(0) ' the first file is at index zero
    If myRARchive.LastError = 0 And item <> Nil Then
      Dim f As FolderItem = GetSaveFolderItem("", item.FileName) ' prompt user for a save location
      If f <> Nil Then
        If myRARchive.ExtractItem(0, f) Then
          MsgBox(item.FileName + " extracted to " + f.AbsolutePath)
        Else
          MsgBox("RAR error " + Str(myRARchive.LastError) + ": " + UnRAR.FormatError(myRARchive.LastError))
        End If
      End If
    End If

