if WScript.arguments.count <= 0 then
    MsgBox "Usage: mailfiles filename..., or drag files onto shortcut"
    WScript.quit 0
end if

const cdoSendUsingPort  = 2
const cdoAnonymous      = 2

sender      = "myname@company.com"
recipient   = "somebody@company.com"
relayserver = "mail.mycompany.com"

set msg = CreateObject("CDO.Message")
set conf = CreateObject("CDO.Configuration")
set msg.configuration = conf

with msg
    .to         = recipient
    .from       = sender
    .subject    = "Files of you"
    .textBody   = "Attached to this message are files of you."

    nfiles = 0
    for each arg in WScript.arguments
        .AddAttachment arg
        nfiles = nfiles + 1
    next
end with

on error resume next
msg.send
send_errno = err.Number
on error goto 0

if send_errno <> 0 then
    MsgBox "Error sending message"
else
    if nfiles = 1 then plural = "" else plural = "s"
    MsgBox "Sent " & nfiles & " file" & plural & " to " & recipient
end if
