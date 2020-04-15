#
#
# Module manifest for module 'ExchangeOnlineShell'
#
# Generated by: Andy Svintsitsky
#
# Generated on: 08/23/2019
#
#

@{

# Script module or binary module file associated with this manifest.
 RootModule = 'ExchangeOnlineShell.psm1'

# Version number of this module.
ModuleVersion = '2.0.3.3'

# Supported PSEditions
# CompatiblePSEditions = @()

# ID used to uniquely identify this module
GUID = 'b5b96354-3688-4885-b01a-4d603ab056b6'

# Author of this module
Author = 'Andy Svintsitsky'

# Company or vendor of this module
CompanyName = 'andysvints.com'

# Copyright statement for this module
Copyright = '(c) 2019 Andy Svintsitsky. All rights reserved.'

# Description of the functionality provided by this module
Description = 'Module for creation a session to manage Exchange Online Shell with or without Proxy Settings.Supports MFA'

# Minimum version of the Windows PowerShell engine required by this module
PowerShellVersion = '5.0'

# Name of the Windows PowerShell host required by this module
# PowerShellHostName = ''

# Minimum version of the Windows PowerShell host required by this module
# PowerShellHostVersion = ''

# Minimum version of Microsoft .NET Framework required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
# DotNetFrameworkVersion = ''

# Minimum version of the common language runtime (CLR) required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
# CLRVersion = ''

# Processor architecture (None, X86, Amd64) required by this module
# ProcessorArchitecture = ''

# Modules that must be imported into the global environment prior to importing this module
# RequiredModules = 

# Assemblies that must be loaded prior to importing this module
RequiredAssemblies = 'Microsoft.Exchange.Management.ExoPowershellModule.dll', 
               'Microsoft.IdentityModel.Clients.ActiveDirectory.dll',
               'Microsoft.IdentityModel.Clients.ActiveDirectory.WindowsForms.dll'

# Script files (.ps1) that are run in the caller's environment prior to importing this module.
# ScriptsToProcess = @()

# Type files (.ps1xml) to be loaded when importing this module
# TypesToProcess = @()

# Format files (.ps1xml) to be loaded when importing this module
# FormatsToProcess = @()

# Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
 NestedModules = '.\Microsoft.Exchange.Management.ExoPowershellModule.dll', 
               '.\Microsoft.IdentityModel.Clients.ActiveDirectory.dll',
               '.\Microsoft.IdentityModel.Clients.ActiveDirectory.WindowsForms.dll'

# Functions to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no functions to export.
FunctionsToExport = 'Connect-ExchangeOnlineShell','Disconnect-ExchangeOnlineShell','Show-EOShellSession'

# Cmdlets to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no cmdlets to export.
CmdletsToExport = 'Connect-ExchangeOnlineShell','Disconnect-ExchangeOnlineShell','Show-EOShellSession'

# Variables to export from this module
#VariablesToExport = '*'

# Aliases to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no aliases to export.
AliasesToExport = '*'

# DSC resources to export from this module
# DscResourcesToExport = @()

# List of all modules packaged with this module
# ModuleList = @()

# List of all files packaged with this module
# FileList = @()

# Private data to pass to the module specified in RootModule/ModuleToProcess. This may also contain a PSData hashtable with additional module metadata used by PowerShell.
PrivateData = @{

    PSData = @{

        # Tags applied to this module. These help with module discovery in online galleries.
         Tags = @( 'Office365', 'ExchangeOnline', 'ExchangeOnlineShell', "O365","Proxy", "ProxySettings", "PSSession", "Session", "PowershellMFA", "ExchangeOnlineShellMFA", "MFA", "MFASupported" )

        # A URL to the license for this module.
        # LicenseUri = ''

        # A URL to the main website for this project.
         ProjectUri = 'https://github.com/andysvints/ExchangeOnlineShell'

        # A URL to an icon representing this module.
        # IconUri = ''

        # ReleaseNotes of this module
         ReleaseNotes = 'Added changes suggested by Taras Fedus @jarlaxle90 - Added IdleSessionTimeout parameter to Connect-ExchangeOnlineShell cmdlet which allows to customize session timeout.'

    } # End of PSData hashtable

} # End of PrivateData hashtable

# HelpInfo URI of this module
# HelpInfoURI = ''



}

