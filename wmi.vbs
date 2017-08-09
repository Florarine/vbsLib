set ns = GetObject("winmgmts:\\.\root\CIMV2")
set disks = ns.InstanceOf("Win32_LogicalDisk")
Set loc = CreateObject("WbemScripting.SWbemLocator")
set svcs = loc.ConnectServer("Java", "root\CIMV2")
ConnectServer(servername, namespace[, username, password], [locale], [authority], [Security], [namedvalueset])

'attributes
SWbemServices.Security_
'methods
SWbemServices.Delete(path)
SWbemServices.ExecMethod(path, methodname, inparams)
SWbemServices.ExecQuery(query)
SWbemServices.InstanceOf(class)

select propertylist from class [where conditions]

'attributes
SWbemObject.Methods_
SWbemObject.Path_
SWbemObject.Properties_
'mathods
SWbemObject.Delete_
SWbemObject.Instances_
SWbemObject.Put_

set obj = GetObject("winmgmts:{impersonationlevel=Impersonate}!" & "//./root/CIMV2:Win32_ComputerSystem")

wscript.echo obj.path_ & vbCRLF
wscript.echo "Properties:"
for each prop in obj.Properties_
    wscript.echo "  ", prop.name
next

wscript.echo
wscript.echo "Methods:"
for each meth in obj.Methods_
    arglist = ""
    for each arg in meth.InParameters.Properties_
        if arglist <> "" then arglist = arglist & ", "
        arglist = arglist & arg.name
    next
    for each arg in meth.OutParameters.Properties_
        if arg.name <> "ReturnValue" then
            if arglist <> "" then arglist = arglist  & ", "
            arglist = arglist & arg.name
        end if
    next
    wscript.echo "  ", meth.name & "(" & arglist & ")"
next
