'attributes
Message.Attachments
Message.AutoGenerateTextBody
Message.BCC
Message.BodyPart
Message.CC
Message.Configuration
Message.DSNOptions
Message.Fields
Message.From
Message.HTMLBody
Message.HTMLBodyPart
Message.MDNRequested
Message.MIMEFormated
Message.Organization
Message.ReplyTo
Message.Sender
Message.Subject
Message.TextBody
Message.TextBodyPart
Message.To
'methods
Message.AddAttachment(path[, username, password])
Message.AddRelatedBodyPart(path, reference, referencetype[, username, password])
Message.CreateMHTMLBody(path[, flags[, username, password]])
Message.Send

cdoDSNOptions
    cdoDSNDefault   0
    cdoDSNNever     1
    cdoDSNFailure   2
    cdoDSNSuccess   4
    cdoDSNDelay     8
    cdoDSNSuccessFailOrDelay    14

cdoReferenceType
    cdoRefTypeId        0
    cdoRefTypeLOcation  1

cdoMHTMLFlags
    CdoSupressNone          0
    CdoSupressImages        1
    CdoSupressBGSounds      2
    CdoSupressFrames        4
    CdoSupressObjects       8
    CdoSupressStyleSheets   16
    CdoSupressAll           31

Fields.Count
Fields.Item(index)
Fields.Update
URN:SCHEMAS:HTTPMAIL:
    content-disposition
    cotent-media-type
    Date
    Fromemail
    Fromname
    Htmldescription
    Importance
    Normalizedsubject
    Priority
    Senderemail
    Sendername
    Textdescription
URN:SCHEMAS:MAILHEADER:
    content-base

cdoImportance
    cdoLow      0
    cdoNormal   1
    cdoHigh     2

cdoPriority
    cdoPriorityNonUrgent
    cdoPriorityNormal
    cdoPriorityUrgent

BodyParts.Count
BodyParts.Item(n)
BodyParts.Add([n])
BodyParts.Delete(x)
BodyParts.DeleteAll

BodyPart.BodyParts
BodyPart.Charset
BodyPart.ContentMediaType
BodyPart.ContentTransferEncoding
BodyPart.Fields
BodyPart.Filename
BodyPart.Parent
BodyPart.AddBodyPart(n)
BodyPart.GetEncodedContentStream()
BodyPart.SaveToFile filename
URN:SCHEMAS:HTTPMAIL:
    Attachmentfilename
    content-disposition-type
URN:SCHEMAS:MAILHEADER:
    content-description
    content-disposition
    content-id
    content-language
    content-location
    content-type
HTTP://SCHEMAS.MICROSOFT.COM/
    Sensitivity

cdoEncodingType
    cdo7bit
    cdo8bit
    cdoBase64
    cdoBinary
    cdoMacBinHex40
    cdoQuotedPrintable
    cdoUuencode

cdoSensitivityValues
    cdoSensitivityNone      0
    cdoPersonal             1
    cdoPrivate              2
    cdoCompanyConfidential  3

Configuration.Fields
Configuration.Load loadFrom

HTTP://SCHEMAS.MICROSOFT.COM/CDO/CONFIGURATION/
    Autopromotebodyparts
    Httpcookies
    Languagecode
    Sendemailaddress
    Sendpassword
    Sendusername
    Senduserreplymailaddress
    Sendusing
    Smtpaccountname
    Smtpauthenticate
    Smtpconnectiontimeout
    Smtpserver
    Smtpserverpickupdirectory
    Smtpserverport
    Smtpusessl
    Urlgetlatestversion
    Urlproxyserver
    Urlproxybypass
URN:SCHEMAS:CALENDAR:
    Timezoneid

cdoConfigSource
    cdoSourceDefaults       -1
    cdoSourceIIS            1
    cdoSourceOutlookExpress 2
cdoProtocolsAuthenticate
    cdoAnonymouse   0
    cdoBasic        1
    cdoNTLM         2
cdoSendUsing
    cdoSendUsingPickup  1
    cdoSendUsingPort    2
cdoTimeZoneId
    cdoUTC
    cdoGMT
    cdoInvalidTimeZone

set msg = CreateObject("CDO.Message")
set config = CreateObject("CDO.Configuration")
set msg.Configuration = config
