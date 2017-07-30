set fso = CreateObject("Scripting.FileSystemObject")
fso.Drives
fso.BuildPath(path, name)
fso.CopyFile source, destination[, overwrite]
fso.CopyFolder source. destination[, overwrite]
fso.CreateFolder(foldername)
fso.CreateTextFile(filename[, overwrite[, unicode]])
fso.DeleteFile filespec[, force]
fso.DeleteFolder folderspec[, force]
fso.DriveExists(drive)
fso.FolderExists(folder)
fso.FileExists(filename)
fso.GetAbsolutePathName(pathspec)
fso.GetFileVersion(filepath)
fso.GetBaseName(filepath)
fso.GetDrive(drivespec)
fso.GetDriveName(path)
fso.GetExtensionName(path)
fso.GetFile(filespec)
fso.GetFileName(filespec)
fso.GetFolder(folderspec)
fso.GetParentFolderName(path)
fso.GetSpecialFolder(folderspec)
'   SpecialFolderConst
const WindowsFolder     = 0
const SystemFolder      = 1
const TemporaryFolder   = 2
fso.GetStandardStream(strm)
'   StandardStreamTypes
const stdIn     = 0
const StdOut    = 1
const StdErr    = 2
fso.GetTempName()
fso.MoveFile source, destination
fso.MoveFolder source, destination
fso.OpenTextFile(filename[, iomode[, create[, format]]])
'   IOMode
const ForReading    = 1
const ForWriting    = 2
const ForAppending  = 8
'   Tristate
Const TristateFalse         = 0
Const TristateTrue          = -1 
Const TristateUseDefault    = -2

sub CreateFullPath(byval path)
    Dim parent
    
    path = fso.GetAbsolutePathName(path)
    parent = fso.GetParentFolderName(path)

    if not fso.FolderExists(parent) then
        CreateFullPath parent
    end if

    if not fso.FolderExists(path) then
        fso.CreateFolder path
    end if
end sub

Drive.AvailableSpace
Drive.DriveLetter
Drive.DriveType
'   DriveTypeConst
const UnknownType   = 0
Const Removable     = 1
Const Fixed         = 2
Const Remote        = 3
Const CDRom         = 4
Const RamDisk       = 5
Drive.FileSystem
Drive.FreeSpace
Drive.IsReady
Drive.Path
Drive.RootFolder
Drive.SerialNumber
Drive.ShareName
Drive.TotalSize
Drive.VolumnName

Folder.Attributes
Folder.DateCreated
Folder.DateLastAccessed
Folder.DateLastModified
Folder.Drive
Folder.Files
Folder.IsRootFolder
Folder.Name
Folder.ParentFolder
Folder.Path
Folder.ShortName
Folder.ShortPath
Folder.Size
Folder.SubFolders
Folder.Type
Folder.Copy destination[, overwrite]
Folder.CreateTextFile(filename[, overwrite[, unicode]])
Folder.Delete [force]
Folder.Move destination
'   FileAttribute
'   Normal      0
'   ReadOnly    1   R/W
'   Hidden      2   R/W
'   System      4   R/W
'   Volumn      8   R/O
'   Directory   16  R/O
'   Achive      32  R/W
'   Alias       64  R/O
'   Compressed  128 R/O

File.Attributes
File.DateCreated
File.DateLastAccessed
File.DateLastModified
File.Drive
File.Name
File.ParentFolder
File.Path
File.ShortName
File.ShortPath
File.Size
File.Type
File.Copy destination[, overwrite]
File.Delete force
File.Move destination
File.OpenAsTextStream([iomode[, format]])

'TextStream
set stream = fso.OpenTextFile(filespec)
set stream = file.OpenAsTextStream(ForReading)
set stream = fso.CreateTextFile(filespec,True)
set stream = file.OpenAsTextStream(ForWriting)
set stream = fso.OpenAsTextStream(filespec, ForAppending, True)
set stream = file.OpenAsTextStream(ForAppending)
'using Unicode
set stream = fso.OpenTextFile(filespec, ForReading, False, TristateTrue)
set stream = fso.CreateTextFile(filespec, True, True)
set stream = fso.OpenAsTextStream(filespec, ForAppending, True, TristateTrue)

TextStream.AtEndOfLine
TextStream.AtEndOfStream
TextStream.Column
TextStream.Line
TextStream.Close
TextStream.Read(nread)
TextStream.ReadAll
TextStream.ReadLine
TextStream.Skip nskip
TextStream.SkipLine
TextStream.Write string
TextStream.WriteBlankLines nlines
TextStream.WriteLine [string]

Wscript.StdIn
Wscript.Stdou
Wscript.StdErr

function binaryval(byref str, offset, n)
    dim i
    binaryval = 0
    for i = offset+n-1 to offset step -1
        binaryval = binaryval*256 + asc(mid(str,i+1,1))
    next
end function
