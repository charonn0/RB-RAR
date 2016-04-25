This project implements bindings for RealStudio and Xojo Win32 applications to extract RAR archives 
using the [official unrar library](http://www.rarlab.com/rar/UnRARDLL.exe). The `unrar.dll` file 
must be stored on Windows' search path or in the same directory as your executable.

The project consists of a [module](https://github.com/charonn0/RB-RAR/wiki/Unrar-module) containing three classes: [Iterator](https://github.com/charonn0/RB-RAR/wiki/UnRAR.Iterator), [IteratorEx](https://github.com/charonn0/RB-RAR/wiki/UnRAR.IteratorEx), and [ArchiveEntry](https://github.com/charonn0/RB-RAR/wiki/UnRAR.ArchiveEntry).

The Iterator class implements the basic, older UnRAR interface. It does not support Unicode filenames, archives with encrypted headers, or extracting into memory. For almost all archives you should prefer the IteratorEx class.

##Examples
###Extract an entire archive
```vbnet
  Dim archive As FolderItem ' assume a valid RAR archive
  Dim outputdir As FolderItem ' assume a valid directory
  Dim rar As New UnRAR.IteratorEx(archive)
  Do Until Not rar.MoveNext(UnRAR.RAR_EXTRACT, outputdir)
    App.YieldToNextThread
  Loop
  If rar.LastError <> UnRAR.ErrorEndArchive Then Raise New UnRAR.RARException(rar.LastError)
  rar.Close
```
###Extract a single file
```vbnet
  Dim index As Integer = 2 ' the third file in the archive
  Do Until rar.CurrentItem.Index = index - 1
    App.YieldToNextThread
  Loop Until Not rar.MoveNext(UnRAR.RAR_SKIP)
  If Not rar.MoveNext(UnRAR.RAR_EXTRACT, outputdir) Then Raise New UnRAR.RARException(rar.LastError)
  rar.Close
```
###Test an entire archive
```vbnet
  Dim archive As FolderItem ' assume a valid RAR archive
  Dim rar As New UnRAR.IteratorEx(archive)
  Do Until Not rar.MoveNext(UnRAR.RAR_TEST)
    App.YieldToNextThread
  Loop
  If rar.LastError <> UnRAR.ErrorEndArchive Then Raise New UnRAR.RARException(rar.LastError)
  rar.Close
```
###Test a single file
```vbnet
  Dim archive As FolderItem ' assume a valid RAR archive
  Dim rar As New UnRAR.IteratorEx(archive)
  Dim index As Integer = 2 ' the third file in the archive
   Do Until rar.CurrentItem.Index = index - 1
    App.YieldToNextThread
  Loop Until Not rar.MoveNext(UnRAR.RAR_SKIP)
  If Not rar.MoveNext(UnRAR.RAR_TEST) And rar.LastError <> UnRAR.ErrorEndArchive Then Raise New UnRAR.RARException(rar.LastError)
  rar.Close
```
