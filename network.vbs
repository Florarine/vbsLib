'attributes
Network.ComputerName
Network.UserDomain
Network.UserName
'methods
Network.AddPrinterConnection LocalName, RemoteName[, UpdateProfile[, UserName, Password]]
Network.AddWindowsPrinterConnection PrinterPath[, DriveName[, Port]]
Network.EnumNetworkDrives()
Network.EnumPrinterConnections()
Network.MapNetworkDrive LocalName, RemoteName[, UpdateProfile[, UserName, Password]]
Network.RemoveNetworkDrive Name[, Forece[, UpdateProfile]]
Network.RemovePrinterConnection Name[, Force[, UpdateProfile]]
Network.SetDefaultPrinter PrinterName

function IsDriveMapped(byval drive)
    drive = ucase(left(drive, 1))
    IsDriveMapped = False
    if not fso.DriveExists(drive) then exit function
    IsDriveMapped = (fso.getDrive(drive).driveType = 3)
end function

function GetDriveMapping(byval drive)
    dim i, maps
    drive = ucase(left(drive, 1))
    set maps = wshNetwork.EnumNetworkDrives
    GetDriveMapping = ""
    for i = 0 to maps.Length-2 step 2
        if ucase(left(maps.item(i), 1)) = drive then
            GetDriveMapping = ucase(maps.item(i+1))
            exit for
        end if
    next
end function

function UnMap(byval drive)
    UnMap = True
    if len(drive) = 1 then drive = drive & ":"
    on error resume next
    err.Clear
    wshNetwork.RemoveNetworkDrive drive, False, True
    if err.Number <> 0 then UnMap = False
    on error goto 0
end function

function MapDrive(byval drive, byval path)
    MapDrive = True
    if len(drive) = 1 then drive = drive & ":"
    if IsDriveMapped(drive) then
        if GetDriveMapping(drive) = ucase(path) then exit function
        UnMap drive
    end if

    err.Clear
    on error resume next
    wshNetwork.MapNetworkDrive drive, path
    if err.Number <> 0 then MapDrive = False
    on error goto 0
end function

function UsePrinter(path)
    UsePrinter = True
    err.Clear
    on error resume next
    wshNetwork.AddWindowsPrinterConnection path
    if err.Number <> 0 then UsePrinter = False
    on error goto 0
end function

function DeletePrinter(byval path)
    dim maps, i, ismapped

    path = ucase(path)
    DeletePrinter = True
    ismapped = False
    set maps = wshNetwork.EnumPrinterConnections
    for i = 0 to maps.Length-2 step 2
        if ucase(maps.item(i+1)) = path then
            ismapped = True
            exit for
        end if
    next
    if not ismapped then exit function

    err.Clear
    on error resume next
    wshNetwork.RemovePrinterConnection path, True, True
    if err.Number <> 0 then DeletePrinter = False
    on error goto 0
end function

function RedirectPrinter(byval device, byval path, byval updateProfile)
    dim maps, i, ismapped
    dim devcolon, devnocolon

    device = ucase(device)
    i = instr(device, ";")
    if i > 0 then
        devcolon = device
        devnocolon = left(device, 1)
    else
        devcolon = device & ":"
        devnocolon = device
    end if
    RedirectPrinter = True
    ismapped = False
    set maps = wshNetwork.EnumPrinterConnections
    for i = 0 to maps.Length-2 step 2
        if maps.item(i) = devnocolon and left(maps.item(i+1), 2) = "\\" then
            ismapped = True
            exit for
        elseif maps.item(i) = devnocolon then
            RedirectPrinter = False
            exit function
        end if
    next
    if ismapped then
        err.Clear
        on error resume next
        wshNetwork.RemovePrinterConnection device, True, updateProfile
        if err.Number <> 0 then
            RedirectPrinter = False
            exit function
        end if
        on error goto 0
    end if
    err.Clear
    on error resume next
    wshNetwork.AddPrinterConnection device, path, updateProfile
    if err.Number <> 0 then RedirectPrinter = False
    on error goto 0
end function
