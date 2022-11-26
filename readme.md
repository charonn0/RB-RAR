## Introduction
**UnRAR** is the official library from RARLabs for extracting [WinRAR archives](http://www.rarlab.com/rar_archiver.htm). This project implements UnRAR bindings for RealStudio and Xojo Win32 applications. Creating archives is not supported by the library.

## Synopsis
The project consists of a [module](https://github.com/charonn0/RB-RAR/wiki/Unrar-module) containing three classes: [Iterator](https://github.com/charonn0/RB-RAR/wiki/UnRAR.Iterator), [IteratorEx](https://github.com/charonn0/RB-RAR/wiki/UnRAR.IteratorEx), and [ArchiveEntry](https://github.com/charonn0/RB-RAR/wiki/UnRAR.ArchiveEntry).

The `Iterator` class implements the basic, older UnRAR interface. It does not support Unicode filenames, archives with encrypted headers, or extracting into memory. For almost all archives you should prefer the `IteratorEx` class.

## Example
### This example extracts an entire RAR archive into a directory
```realbasic
  Dim archive As FolderItem ' assume a valid RAR archive
  Dim outputdir As FolderItem ' assume a valid directory
  Dim rar As New UnRAR.IteratorEx(archive)
  Do
    If Not rar.MoveNext(UnRAR.RAR_EXTRACT, outputdir) Then Exit Do
  Loop
  rar.Close
```

## How to incorporate RB-RAR into your Realbasic/Xojo project
### Import the UnRAR module
1. Download the RB-RAR project either in [ZIP archive format](https://github.com/charonn0/RB-RAR/archive/master.zip) or by cloning the repository with your Git client.
2. Open the RB-RAR project in REALstudio or Xojo. Open your project in a separate window.
3. Copy the UnRAR module into your project and save.

### Ensure the UnRAR.dll shared library is installed
You will need to ship the UnRAR.dll file with your application and install it into the same directory as your app. A pre-built Win32 DLL is available [here](http://www.rarlab.com/rar/UnRARDLL.exe), or you can [build it yourself from source](http://www.rarlab.com/rar/unrarsrc-5.3.11.tar.gz). 

## More examples
### Extract a single file
```realbasic
  Dim index As Integer = 2 ' the third file in the archive
  Dim archive As FolderItem ' assume a valid RAR archive
  Dim outputfile As FolderItem ' assume a valid, non-existing file
  Dim rar As New UnRAR.IteratorEx(archive)
  rar.Reset(index) ' move to the index
  Call rar.MoveNext(UnRAR.RAR_EXTRACT, outputfile)
  rar.Close
```
### Test an entire archive
```realbasic
  Dim archive As FolderItem ' assume a valid RAR archive
  Dim rar As New UnRAR.IteratorEx(archive)
  Do
    If Not rar.MoveNext(UnRAR.RAR_TEST) Then Exit Do
  Loop
  If rar.LastError <> UnRAR.ErrorEndArchive Then 
    MsgBox("Test failed!")
  End If
  rar.Close
```
### Test a single file
```realbasic
  Dim index As Integer = 2 ' the third file in the archive
  Dim archive As FolderItem ' assume a valid RAR archive
  Dim rar As New UnRAR.IteratorEx(archive)
  rar.Reset(index) ' move to the index  
  If Not rar.MoveNext(UnRAR.RAR_TEST) And rar.LastError <> UnRAR.ErrorEndArchive Then 
    MsgBox("Test failed!")
  End If
  rar.Close
```
