@{

# Script module or binary module file associated with this manifest.
RootModule = 'Invoke-ConnectExchange.psm1'

# Version number of this module.
ModuleVersion = '1.1.0'

# ID used to uniquely identify this module
GUID = '5b658ad8-d1f9-4ab2-8505-89bc7c8615b0'

# Author of this module
Author = 'thomas.illiet'

# Company or vendor of this module
CompanyName = 'netboot.fr'

# Copyright statement for this module
Copyright = '(c) 2018 Netboot. All rights reserved.'

# Description of the functionality provided by this module
Description = 'This module will allow you to connect with Exchange using PowerShell.'

# Functions to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no functions to export.
FunctionsToExport = @('Test-ExchangeSession', 'Invoke-ConnectExchange')

# Cmdlets to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no cmdlets to export.
CmdletsToExport = ''

# Variables to export from this module
VariablesToExport = ''

# Aliases to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no aliases to export.
AliasesToExport = ''

# Private data to pass to the module specified in RootModule/ModuleToProcess. This may also contain a PSData hashtable with additional module metadata used by PowerShell.
PrivateData = @{

    PSData = @{

        # Tags applied to this module. These help with module discovery in online galleries.
        Tags = @('Office365', 'Exchange', 'Connect')

        # A URL to the license for this module.
        LicenseUri = 'https://raw.githubusercontent.com/Netboot-France/Invoke-ConnectExchange/master/LICENSE'

        # A URL to the main website for this project.
        ProjectUri = 'https://github.com/Netboot-France/Invoke-ConnectExchange'

        # A URL to an icon representing this module.
        IconUri = 'https://raw.githubusercontent.com/Netboot-France/Invoke-ConnectExchange/master/Resource/Icon.png'

    } # End of PSData hashtable

} # End of PrivateData hashtable

}

