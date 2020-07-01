# NIC-Configuration-Bot
An idempotent automation script for configuring network interfaces on Windows clients.

This script offers a solution to the problem of standardizing NIC
configuration settings in a population of computers with heterogeneous NIC
makes and models.  The problem being that there is no standard for the
naming of configuration settings or their option enumerations.  This solution
resolves it by automating the survey of all NIC makes and models in the
environment, providing a configuration file for specifying desired settings
by NIC type, and then applying those configuration specifications back to the
environment.  The script should be used in three (3) phases.

Phase 1: Discovery - 
When run with "/mode:discover", this script will save to the XML file
specified in "/nics:[NICSFILEPATH]" the full set of driver configuration
options for every make and model of active network adapter found in the list
of computers specified in "/computers:[COMPUTERFILEPATH]".  If the nics file
exists, it will be loaded.  If the loaded file is not valid XML or the XML's
root node is not "<nics>", the script will create a new XML structure and
replace the file.  Each NIC make and model will only be written to the XML
once, keying on NIC description as found in target computer registries.

Phase 2: Design - 
In "Phase 1", the script inserts an attribute called "push" into each
"<param>" tag of the generated XML.  This attribute allows users to specify
a value for each parameter and NIC.  Users must pay attention to the other
parameter attributes when setting this value to ensure that "push" values
are valid for their parameters.  For example, a "<param>" with attribute
"type" set to "int" may also have attributes "min" and "max".  These will
guide users in setting an integer value for the "push" that falls within a
proper range.  Many parameters, however, are of "type=enum".  Enum types
have a specific set of valid values they can be set to.  In these cases, a
"<param>" node will have an "<enum>" child node and "<enums>" grand-
children.  These will guide users on what values may be used to "push"
to an "enum" type "<param>" and what those values mean.

Phase 3:  Implement - 
Once the XML has been updated with the desired "push" values for each
"<nic>" and "<param>", the script can be run with "/mode:implement" to apply
the values specified in "/nics:[NICSFILEPATH]" to the list of computers in
"/computers:[COMPUTERSFILEPATH]".  If the "nics" file cannot be loaded or
the XML structure is not as expected, the script will abort.

The last step per computer is for the script to read the value of
"constRebootTargets" and if its true, reboot the target computer.

Once the script has validated the arguments and started processing, it
will offer some information at the console.  First, a header block with the
script runtime configuration will be shown.  Then, a dynamically updating
progress indicator will show the computer name currently processing and it's
(X of Y) position in the total queue.



Arguments:	/computers:[COMPUTERSFILEPATH]
				path to a text file that lists target hostnames

			/nics:[NICSFILEPATH]
				path to save the driver XML database to

			/mode:[DISCOVER|IMPLEMENT]

Requires:	Executor system OS WIndows XP Professional of above
			Executor account has "Administrators" level access to target
				computers
