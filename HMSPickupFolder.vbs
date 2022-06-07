const PICKUPFOLDER = "C:\Program Files (x86)\hMailServer\Pickup"
const INTERVAL = 1
const TIMEOUT = 10000
const WMI_TIMEOUT = -2147209215
set oWMI = GetObject("winmgmts://./root/cimv2")
set oFS = CreateObject("Scripting.FileSystemObject")
set oMonitoredFolder = CreateMonitoredFolder(PICKUPFOLDER)
do
	on error resume next
	set oFolderEvent = oMonitoredFolder.NextEvent(TIMEOUT)
	if err.number = WMI_TIMEOUT then
		on error goto 0
	else
		on error goto 0
		select case oFolderEvent.Path_.Class
			case "__InstanceCreationEvent": DispatchFolderChangeEvent oFolderEvent.TargetInstance
			'case "__InstanceModificationEvent" : 
			'case "__InstanceDeletionEvent"     : 
		end select
	end if
loop

function CreateMonitoredFolder(sFolderPath)
	aPath = Split(oFS.GetAbsolutePathName(sFolderPath), ":")
	sDrive  = aPath(0) & ":"
	sDirectory  = Replace(aPath(1), "\", "\\")
	if Right(sDirectory, 2) <> "\\" then 
		sDirectory = sDirectory & "\\"
	end if
	sWMIQuery = "SELECT * FROM __InstanceOperationEvent WITHIN " & INTERVAL & _
		  " WHERE Targetinstance ISA 'CIM_DataFile' AND TargetInstance.Drive='" & sDrive & "'" & _
		  " AND TargetInstance.Path='" & sDirectory & "'"
	Set CreateMonitoredFolder = oWMI.ExecNotificationQuery(sWMIQuery)
End Function

sub DispatchFolderChangeEvent(oInstance)
	set oFolder = oFS.GetFolder(PICKUPFOLDER)
	for each oFile in oFolder.Files
		wscript.echo FormatDateTime(Now, vbGeneralDate) & "  Processing file " & oFile.name
		PickupOutgoingMailFile(oFile)
	next
end sub

sub PickupOutgoingMailFile(oFile)
	on error resume next
	set oMailMessage = CreateObject("hMailServer.Message")
	oFS.CopyFile oFile.Path, oMailMessage.FileName, true
	oMailMessage.RefreshContent
	sOriginalTo = oMailMessage.HeaderValue("To")
	sOriginalCC = oMailMessage.HeaderValue("CC")
	oMailMessage.ClearRecipients
	wscript.echo FormatDateTime(Now, vbGeneralDate) & "  Setting MAIL FROM: to " & CleanAddress(oMailMessage.HeaderValue("From"))
	oMailMessage.FromAddress = CleanAddress(oMailMessage.HeaderValue("From"))
	for each sRecipient in SplitBetween(sOriginalTo, "<", ">") 
		wscript.echo FormatDateTime(Now, vbGeneralDate) & "  Adding To: recipient " & sRecipient
		oMailMessage.AddRecipient "", CleanAddress(sRecipient)
	next
	for each sRecipient in SplitBetween(sOriginalCC, "<", ">") 
		wscript.echo FormatDateTime(Now, vbGeneralDate) & "  Adding CC: recipient " & sRecipient		
		oMailMessage.AddRecipient "", CleanAddress(sRecipient)
	next
	oMailMessage.HeaderValue("To") = sOriginalTo
	oMailMessage.HeaderValue("CC") = sOriginalCC
	wscript.echo FormatDateTime(Now, vbGeneralDate) & "  File sent."
	oMailMessage.Save
	oFS.DeleteFile oFile.path, true
	on error goto 0
end sub

function SplitBetween(sString, sFrom, sTo)
	SPLITCHAR = Chr(10)
	sRemainder = sString
	do while InStr(sRemainder, sFrom) > 0
		sRemainder = Mid(sRemainder, InStr(sRemainder, sFrom) + Len(sFrom))
		if InStr(sRemainder, sTo) > 0 then
			if sBetween = "" then 
				sBetween = Mid(sRemainder, 1, InStr(sRemainder, sTo) - 1)
			else
				sBetween = sBetween & SPLITCHAR & Mid(sRemainder, 1, InStr(sRemainder, sTo) - 1)
			end if
			sRemainder = Mid(sRemainder, InStr(sRemainder, sTo))
		end if
	loop
	SplitBetween = Split(sBetween, SPLITCHAR)
end function

function CleanAddress(sAddress)
	dim i
	i = InStrRev(sAddress, "<")
	if i > 0 then
		sAddress = Mid(sAddress, i + 1)
		i = InStr(sAddress, ">")
		if i > 0 then
			sAddress = Mid(sAddress, 1, i - 1)
		end if
		sAddress = CleanAddress(sAddress)
	end if
	CleanAddress = lcase(sAddress)
end function
