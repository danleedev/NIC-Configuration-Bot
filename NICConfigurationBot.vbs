option explicit
'******************************************************************************
const constTitle = "NIC Configuration Bot"
const constVersion = "1.0"
const constUsage = "/computers:[COMPUTERSFILEPATH] /nics:[NICSFILEPATH] /mode:[DISCOVER|IMPLEMENT]"
'*
'* This script offers a solution to the problem of standardizing NIC
'* configuration settings in a population of computers with heterogeneous NIC
'* makes and models.  The problem being that there is no standard for the
'* naming of configuration settings or their option enumerations.  This solution
'* resolves it by automating the survey of all NIC makes and models in the
'* environment, providing a configuration file for specifying desired settings
'* by NIC type, and then applying those configuration specifications back to the
'* environment.  The script should be used in three (3) phases.
'*
'* Phase 1: Discovery
'* When run with "/mode:discover", this script will save to the XML file
'* specified in "/nics:[NICSFILEPATH]" the full set of driver configuration
'* options for every make and model of active network adapter found in the list
'* of computers specified in "/computers:[COMPUTERFILEPATH]".  If the nics file
'* exists, it will be loaded.  If the loaded file is not valid XML or the XML's
'* root node is not "<nics>", the script will create a new XML structure and
'* replace the file.  Each NIC make and model will only be written to the XML
'* once, keying on NIC description as found in target computer registries.
'*
'* Phase 2: Design
'* In "Phase 1", the script inserts an attribute called "push" into each
'* "<param>" tag of the generated XML.  This attribute allows users to specify
'* a value for each parameter and NIC.  Users must pay attention to the other
'* parameter attributes when setting this value to ensure that "push" values
'* are valid for their parameters.  For example, a "<param>" with attribute
'* "type" set to "int" may also have attributes "min" and "max".  These will
'* guide users in setting an integer value for the "push" that falls within a
'* proper range.  Many parameters, however, are of "type=enum".  Enum types
'* have a specific set of valid values they can be set to.  In these cases, a
'* "<param>" node will have an "<enum>" child node and "<enums>" grand-
'* children.  These will guide users on what values may be used to "push"
'* to an "enum" type "<param>" and what those values mean.
'*
'* Phase 3:  Implement
'* Once the XML has been updated with the desired "push" values for each
'* "<nic>" and "<param>", the script can be run with "/mode:implement" to apply
'* the values specified in "/nics:[NICSFILEPATH]" to the list of computers in
'* "/computers:[COMPUTERSFILEPATH]".  If the "nics" file cannot be loaded or
'* the XML structure is not as expected, the script will abort.
'*
'* The last step per computer is for the script to read the value of
'* "constRebootTargets" and if its true, reboot the target computer.
'*
'* Once the script has validated the arguments and started processing, it
'* will offer some information at the console.  First, a header block with the
'* script runtime configuration will be shown.  Then, a dynamically updating
'* progress indicator will show the computer name currently processing and it's
'* (X of Y) position in the total queue.
'*
'*
'*
'* Arguments:	/computers:[COMPUTERSFILEPATH]
'*					path to a text file that lists target hostnames
'*
'*				/nics:[NICSFILEPATH]
'*					path to save the driver XML database to
'*
'*				/mode:[DISCOVER|IMPLEMENT]
'*
'* Requires:	Executor system OS Windows XP Professional or above
'*		Executor account has "Administrators" level access to target
'*		computers
'*
'*
'* Change Log
'*	Verson 1.0 - 2011.12.04
'*	- first executable; implements the initial requirement
'******************************************************************************

wscript.echo constTitle & vbCrLf & "Version " & constVersion & vbCrLf

'Configuration
	const constNICsControlClassRegistryPath = "SYSTEM\CurrentControlSet\Control\Class\{4D36E972-E325-11CE-BFC1-08002BE10318}"
	const constRebootTargets = false

'Definition

	'constants
		'debug levels
			const constDebugLevelOff = 0
			const constDebugLevelConsole = 1
			const constDebugLevelLog = 2
			const constDebugLevelHalt = 3
		
		'registry hives
			const constRegistryHiveHKLM = 2147483650
			const constRegistryHiveHKCR = 2147483648
			const constRegistryHiveHKCC = 2147483653
			const constRegistryHiveHKCU = 2147483649

		'registry value data types
			const constRegistryValueDataTypeREG_SZ = 1
			const constRegistryValueDataTypeREG_EXPAND_SZ = 2
			const constRegistryValueDataTypeREG_BINARY = 3
			const constRegistryValueDataTypeREG_DWORD = 4
			const constRegistryValueDataTypeREG_MULTI_SZ = 7

		'registry access modes
			const constRegistryAccessQuery = 1

		'file modes
			const constFileModeRead = 1
			
		'remote reboot parameters
			const constRebootParameterCooperativeLogoff = 2
			const constRebootParameterForcedLogoff = 6

	'variables
		dim strMode
		dim astrComputers
		dim strComputer
		dim objComputerOS
		dim intComputersCount
		dim intComputerCurrentIndex
		dim strComputerIP
		dim strComputerNICIndex
		dim strComputerNICRegistryKey
		dim strComputerNICRegistryPath
		dim strComputerNICDescription
		dim astrComputerNICParameters
		dim strComputerNICParameter
		dim astrComputerNICParameterAttributes
		dim strComputerNICParameterAttribute
		dim astrComputerNICParameterEnums
		dim strComputerNICParameterEnum
		dim bolComputerNICConfigurationChanged
		dim objXMLDoc
		dim xmlNICs
		dim xmlNIC
		dim xmlNICDescription
		dim xmlNICParameter
		dim xmlNICParameterAttribute
		dim xmlNICParameterEnums
		dim xmlNICParameterEnum
		dim xmlNICParameterEnumValue
		dim xmlNICParameterEnumLabel
		dim bolNICAlreadyDefined


'Validation
	if wscript.arguments.named.count <> 3 then wscript.echo "USAGE: " & wscript.scriptName & " " & constUsage & vbCrLf & "ERROR:  Script requires exactly 3 arguments.  Aborting." : wscript.quit
	if not wscript.arguments.named.exists("computers") then wscript.echo "USAGE: " & wscript.scriptName & " " & constUsage & vbCrLf & "ERROR:  Script requires argument '/computers:[COMPUTERSFILEPATH]'.  Aborting." : wscript.quit
	if not wscript.arguments.named.exists("nics") then wscript.echo "USAGE: " & wscript.scriptName & " " & constUsage & vbCrLf & "ERROR:  Script requires argument '/nics:[NICSFILEPATH]'.  Aborting." : wscript.quit
	if not wscript.arguments.named.exists("mode") then wscript.echo "USAGE: " & wscript.scriptName & " " & constUsage & vbCrLf & "ERROR:  Script requires argument '/mode:[DISCOVER|IMPLEMENT]'.  Aborting." : wscript.quit
	if strcomp ( wscript.arguments.named.item("mode") , "discover" , 1 ) <> 0 and strcomp ( wscript.arguments.named.item("mode") , "implement" , 1 ) <> 0 then wscript.echo "USAGE: " & wscript.scriptName & " " & constUsage & vbCrLf & "ERROR:  Script argument '/mode:' specified as '" & wscript.arguments.named.item("mode") & "', but must be either 'DISCOVER' or 'IMPLEMENT' (case insensitive).  Aborting." : wscript.quit



'Initialization
	strMode = wscript.arguments.named.item("mode")
	astrComputers = astrLoadHostnames(wscript.arguments.named.item("computers"))
	set objXMLDoc = createObject("Microsoft.XMLDOM")

	if createObject("Scripting.FileSystemObject").FileExists(wscript.arguments.named.item("nics")) then
		on error resume next
			objXMLDoc.load(wscript.arguments.named.item("nics"))
		if err.number <> 0 then
			if strComp ( strMode , "implement" , 1 ) then
				wscript.echo "USAGE: " & wscript.scriptName & " " & constUsage & vbCrLf & "ERROR:  File specified '" & wscript.arguments.named.item("nics") & "' exists, but could not be loaded as XML.  This is a fatal error in 'implement' mode.  Aborting." : wscript.quit
			else
				wscript.echo "WARNING:  File specified '" & wscript.arguments.named.item("nics") & "' exists, but could not be loaded as XML.  File will be replaced by a valid XML file with the same filename."
			end if
		end if
		on error goto 0
		if strcomp ( objXMLDoc.documentElement.nodeName , "nics" , 1 ) = 0 then
			set xmlNICs = objXMLDoc.documentElement
		else
			if strComp ( strMode , "implement" , 1 ) then
				wscript.echo "USAGE: " & wscript.scriptName & " " & constUsage & vbCrLf & "ERROR:  File specified '" & wscript.arguments.named.item("nics") & "' exists and is valid XML, but the root node's tagname is not '<nics>'.  This is a fatal error in 'implement' mode.  Aborting." : wscript.quit
			else
				wscript.echo "WARNING:  File specified '" & wscript.arguments.named.item("nics") & "' exists and is valid XML, but the root node's tagname is not '<nics>'.  File will be replaced by an XML file with the same filename and the root node '<nics>'."
				set objXMLDoc = nothing
				set objXMLDoc = createObject("Microsoft.XMLDOM")
				set xmlNICs = objXMLDoc.createElement("nics")
				objXMLDoc.appendChild xmlNICs
			end if
		end if
	else
		if strComp ( strMode , "implement" , 1 ) = 0 then
			wscript.echo "USAGE: " & wscript.scriptName & " " & constUsage & vbCrLf & "ERROR:  Specified NIC configuration XML file '" & wscript.arguments.named.item("nics") & "' does not exist or is inaccessible.  This is a fatal error in 'implement' mode.  Aborting." : wscript.quit
		else
			set xmlNICs = objXMLDoc.createElement("nics")
			objXMLDoc.appendChild xmlNICs
		end if
	end if

	
'Execution

	wscript.echo "Operating Mode:		" & strMode
	wscript.echo "Computer List:		" & wscript.arguments.named.item("computers")
	wscript.echo "NIC Configuration XML:	" & wscript.arguments.named.item("nics")
	if strComp ( strMode , "implement" , 1 ) = 0 then wscript.echo "Reboot Flag:		" & CStr(constRebootTargets)
	wscript.echo "Computer Count:		" & ubound(astrComputers)

	intComputerCurrentIndex = 0
	for each strComputer in astrComputers

		wscript.stdout.write vbCr & "Progress:		" & intComputerCurrentIndex
		intComputerCurrentIndex = intComputerCurrentIndex + 1
	
		strComputerIP = strGetHostIPAddress ( strComputer )
		if not isEmpty ( strComputerIP ) then
			strComputerNICIndex = strGetNICIndex ( strComputerIP )
			if not isNull ( strComputerNICIndex ) then

				strComputerNICRegistryKey = strComputerNICIndex
				while len(strComputerNICRegistryKey) < 4
					strComputerNICRegistryKey = "0" & strComputerNICRegistryKey
				wend
				strComputerNICRegistryPath = constNICsControlClassRegistryPath & "\" & strComputerNICRegistryKey

				strComputerNICDescription = varGetRegistryValue ( strComputer , constRegistryHiveHKLM , strComputerNICRegistryPath , "driverdesc" )
				if not isNull ( strComputerNICDescription ) then

					if strComp ( strMode , "implement" , 1 ) = 0 then
						for each xmlNIC in xmlNICs.getElementsByTagName("nic")
							if strComp ( xmlNIC.attributes.getNamedItem("description").value , strComputerNICDescription , 1 ) = 0 then
								bolComputerNICConfigurationChanged = false
								for each xmlNICParameter in xmlNIC.getElementsByTagName("param")
									if strComp ( xmlNICParameter.attributes.getNamedItem("push").value , "" , 1 ) <> 0 then
										subSetRegistryValue strComputer , constRegistryHiveHKLM , strComputerNICRegistryPath , xmlNICParameter.attributes.getNamedItem("name").value , xmlNICParameter.attributes.getNamedItem("push").value
										bolComputerNICConfigurationChanged = true
									end if
								next

								if constRebootTargets = true and bolComputerNICConfigurationChanged = true then
									for each objComputerOS in getObject("winmgmts:{impersonationLevel=impersonate,(Debug,Shutdown)}\\" & strComputer & "\root\cimv2").execQuery("Select * from Win32_OperatingSystem")
										objComputerOS.win32Shutdown constRebootParameterForcedLogoff , 0
									next
								end if
							end if
						next
						
					else
					
						bolNICAlreadyDefined = false
						for each xmlNIC in xmlNICs.getElementsByTagName("nic")
							if strComp ( xmlNIC.attributes.getNamedItem("description").value , strComputerNICDescription , 1 ) = 0 then bolNICAlreadyDefined = true
						next
					
						if bolNICAlreadyDefined <> true then

							set xmlNIC = objXMLDoc.createElement("nic")
							set xmlNICDescription = objXMLDoc.createAttribute("description")
							xmlNICDescription.text = strComputerNICDescription
							xmlNIC.attributes.setNamedItem xmlNICDescription

							astrComputerNICParameters = astrGetRegistryKeySubkeyNames ( strComputer , constRegistryHiveHKLM , strComputerNICRegistryPath & "\ndi\params" )
							if not isNull ( astrComputerNICParameters ) then

								for each strComputerNICParameter in astrComputerNICParameters

									set xmlNICParameter = objXMLDoc.createElement("param")

									set xmlNICParameterAttribute = objXMLDoc.createAttribute("name")
									xmlNICParameterAttribute.text = strComputerNICParameter
									xmlNICParameter.attributes.setNamedItem xmlNICParameterAttribute
									
									astrComputerNICParameterAttributes = astrGetRegistryKeyValueNames ( strComputer , constRegistryHiveHKLM , strComputerNICRegistryPath & "\ndi\params\" & strComputerNICParameter )
									if not isNull ( astrComputerNICParameterAttributes ) then
										for each strComputerNICParameterAttribute in astrComputerNICParameterAttributes

											strComputerNICParameterAttribute = lcase ( strComputerNICParameterAttribute )
											set xmlNICParameterAttribute = objXMLDoc.createAttribute(strComputerNICParameterAttribute)
											xmlNICParameterAttribute.text = varGetRegistryValue ( strComputer , constRegistryHiveHKLM , strComputerNICRegistryPath & "\ndi\params\" & strComputerNICParameter , strComputerNICParameterAttribute )
											xmlNICParameter.attributes.setNamedItem xmlNICParameterAttribute

										next
									end if

									set xmlNICParameterAttribute = objXMLDoc.createAttribute("push")
									xmlNICParameterAttribute.text = ""
									xmlNICParameter.attributes.setNamedItem xmlNICParameterAttribute

									if strComp ( xmlNICParameter.attributes.getNamedItem("type").value , "enum" , 1 ) = 0 then
									
										set xmlNICParameterEnums = objXMLDoc.createElement("enums")
										
										astrComputerNICParameterEnums = astrGetRegistryKeyValueNames ( strComputer , constRegistryHiveHKLM , strComputerNICRegistryPath & "\ndi\params\" & strComputerNICParameter & "\enum" )
										for each strComputerNICParameterEnum in astrComputerNICParameterEnums
											
											set xmlNICParameterEnum = objXMLDoc.createElement("enum")
											
											strComputerNICParameterEnum = lcase ( strComputerNICParameterEnum )
											set xmlNICParameterEnumValue = objXMLDoc.createAttribute("value")
											xmlNICParameterEnumValue.text = strComputerNICParameterEnum
											xmlNICParameterEnum.attributes.setNamedItem xmlNICParameterEnumValue
											
											set xmlNICParameterEnumLabel = objXMLDoc.createAttribute("label")
											xmlNICParameterEnumLabel.text = lcase ( varGetRegistryValue ( strComputer , constRegistryHiveHKLM , strComputerNICRegistryPath & "\ndi\params\" & strComputerNICParameter & "\enum" , strComputerNICParameterEnum ) )
											xmlNICParameterEnum.attributes.setNamedItem xmlNICParameterEnumLabel
	
											xmlNICParameterEnums.appendChild(xmlNICParameterEnum)
										next
										xmlNICParameter.appendChild(xmlNICParameterEnums)
									
									end if

									xmlNIC.appendChild(xmlNICParameter)
									set xmlNICParameter = nothing
									
								next

							end if

							xmlNICs.appendChild(xmlNIC)
						end if
					
					objXMLDoc.save(wscript.arguments.named.item("nics"))	
					end if
					
				end if

			end if

		end if

	next
	wscript.echo vbCrLf & vbCrLf & "Done."


'Abstraction

function astrLoadHostnames ( strComputernamesFilePath )
	dim objFileSystem
	dim strComputernamesFileTextRaw

	On Error Resume Next
		set objFileSystem = CreateObject("Scripting.FileSystemObject")
	On Error Goto 0

	if isObject ( objFileSystem ) then
		strComputernamesFileTextRaw = objFileSystem.OpenTextFile(strComputernamesFilePath,constFileModeRead).ReadAll
		strComputernamesFileTextRaw = replace ( strComputernamesFileTextRaw , " " , "" )
		strComputernamesFileTextRaw = replace ( strComputernamesFileTextRaw , vbCrLf , " " )
		do while ( instr ( strComputernamesFileTextRaw , "  " ) > 0 )
			strComputernamesFileTextRaw = replace ( strComputernamesFileTextRaw , "  " , " " )
		loop
		strComputernamesFileTextRaw = trim ( strComputernamesFileTextRaw )
		strComputernamesFileTextRaw = replace ( strComputernamesFileTextRaw , " ", vbCrLf )

		astrLoadHostnames = Split ( strComputernamesFileTextRaw , vbCrLf )
	else
		astrLoadHostnames = null
	end if
	
	set objFileSystem = nothing
end Function

function strGetHostIPAddress ( strComputername )
	dim colPingResults
	dim objPingResult
	dim strComputerIPAddress
	
	set colPingResults = GetObject("winmgmts://./root/cimv2").ExecQuery("select * from win32_pingstatus where address = '" & strComputername & "'")

	for each objPingResult in colPingResults
		strComputerIPAddress = objPingResult.protocoladdress
	next

	if len(strComputerIPAddress) > 0 then
		strGetHostIPAddress = strComputerIPAddress
	else
		strGetHostIPAddress = null
	end if

	set objPingResult = nothing
	set colPingResults = nothing
end function

function strGetNICIndex ( strIPAddress )
	dim colNICs
	dim objNIC
	dim strNICIndex

	On Error Resume Next
		set colNICs = GetObject("winmgmts://" & strIPAddress & "/root/cimv2").ExecQuery("select * from win32_networkadapterconfiguration where ipenabled = true")
	On Error Goto 0

	if isObject ( colNICs ) then
		for each objNIC in colNICs
			if strcomp(objNIC.IPAddress(0),strIPAddress,1) = 0 then
				strNICIndex = objNIC.Index
			end if
		next
		strGetNICIndex = strNICIndex
	else
		strGetNICIndex = null
	end if
	
	set objNIC = nothing
	set colNICs = nothing
end function

function astrGetRegistryKeySubkeyNames ( strComputername , hexHive , strKeyPath )
	dim objRegistry

	On Error Resume Next
		set objRegistry = GetObject("winmgmts://" & strComputername & "/root/default:StdRegProv")
	On Error Goto 0

	if isObject ( objRegistry ) then
		objRegistry.EnumKey hexHive, strKeyPath, astrGetRegistryKeySubkeyNames
	else
		astrGetRegistryKeySubkeyNames = null
	end if
	
	set objRegistry = nothing
end Function

function astrGetRegistryKeyValueNames ( strComputername , hexHive , strKeyPath )
	dim objRegistry
	dim astrEnumeratedValueNames
	dim intLoopIterator

	On Error Resume Next
		set objRegistry = GetObject("winmgmts://" & strComputername & "/root/default:StdRegProv")
	On Error Goto 0

	if isObject ( objRegistry ) then
		objRegistry.EnumValues hexHive, strKeyPath, astrGetRegistryKeyValueNames, null
	else
		astrGetRegistryKeyValueNames = null
	end if

	set objRegistry = nothing
end function

function varGetRegistryValue ( strComputername , hexHive , strKeyPath, strValueName )
	dim objRegistry
	dim astrEnumeratedValueNames
	dim aintEnumeratedValueDataTypes
	dim intLoopIterator
	dim strEnumeratedValueName
	dim intValueDataType

	On Error Resume Next
		set objRegistry = GetObject("winmgmts://" & strComputername & "/root/default:StdRegProv")
	On Error Goto 0

	if isObject ( objRegistry ) then
		objRegistry.EnumValues hexHive, strKeyPath, astrEnumeratedValueNames, aintEnumeratedValueDataTypes

		for intLoopIterator = lbound(astrEnumeratedValuenames) to ubound(astrEnumeratedValueNames)
			if strComp ( astrEnumeratedValueNames(intLoopIterator) , strValueName , 1 ) = 0 then intValueDataType = aintEnumeratedValueDataTypes(intLoopIterator)
		next

		select case intValueDataType
			case constRegistryValueDataTypeREG_SZ objRegistry.GetStringValue hexHive, strKeyPath, strValueName, varGetRegistryValue
			case constRegistryValueDataTypeREG_EXPAND_SZ objRegistry.GetExpandedStringValue hexHive, strKeyPath, strValueName, varGetRegistryValue
			case constRegistryValueDataTypeREG_BINARY objRegistry.GetBinaryValue hexHive, strKeyPath, strValueName, varGetRegistryValue
			case constRegistryValueDataTypeREG_DWORD objRegistry.GetDWORDValue hexHive, strKeyPath, strValueName, varGetRegistryValue
			case constRegistryValueDataTypeREG_MULTI_SZ objRegistry.GetMultiStringValue hexHive, strKeyPath, strValueName, varGetRegistryValue
			case else varGetRegistryValue = ""
		end select
	else
		varGetRegistryValue = ""
	end if

	set objRegistry = nothing
end function

sub subSetRegistryValue ( strComputername , hexHive , strKeyPath, strValueName, varValue )
	dim objRegistry
	dim astrEnumeratedValueNames
	dim aintEnumeratedValueDataTypes
	dim intLoopIterator
	dim strEnumeratedValueName
	dim intValueDataType

	On Error Resume Next
		set objRegistry = GetObject("winmgmts://" & strComputername & "/root/default:StdRegProv")
	On Error Goto 0

	if isObject ( objRegistry ) then
		objRegistry.EnumValues hexHive, strKeyPath, astrEnumeratedValueNames, aintEnumeratedValueDataTypes

		for intLoopIterator = lbound(astrEnumeratedValuenames) to ubound(astrEnumeratedValueNames)
			if strComp ( astrEnumeratedValueNames(intLoopIterator) , strValueName , 1 ) = 0 then intValueDataType = aintEnumeratedValueDataTypes(intLoopIterator)
		next

		select case intValueDataType
			case constRegistryValueDataTypeREG_SZ objRegistry.SetStringValue hexHive, strKeyPath, strValueName, varValue
			case constRegistryValueDataTypeREG_EXPAND_SZ objRegistry.SetExpandedStringValue hexHive, strKeyPath, strValueName, varValue
			case constRegistryValueDataTypeREG_BINARY objRegistry.SetBinaryValue hexHive, strKeyPath, strValueName, varValue
			case constRegistryValueDataTypeREG_DWORD objRegistry.SetDWORDValue hexHive, strKeyPath, strValueName, varValue
			case constRegistryValueDataTypeREG_MULTI_SZ objRegistry.SetMultiStringValue hexHive, strKeyPath, strValueName, varValue
		end select
	end if

	set objRegistry = nothing
end sub

sub subDebug ( strMessage )
	if constDebugLevel >= constDebugLevelConsole then CreateObject("Scripting.FileSystemObject").OpenTextFile( replace ( wscript.scriptname, ".vbs", "" ) & ".log",constFileModeAppend,true).WriteLine ( Now & ":" & strMessage )
	if constDebugLevel >= constDebugLevelLog then wscript.echo Now & ":" & strMessage
	if constDebugLevel = constDebugLevelHalt then wscript.quit
end sub
