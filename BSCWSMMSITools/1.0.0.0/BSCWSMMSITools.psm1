#=================================================================
#region - Description
#=================================================================
# Tools to work with Microsoft Installer files
#=================================================================
#endregion - Description
#=================================================================

#=================================================================
#region - Define and Export Module Variables / Functions
#=================================================================
#=================================================================
#endregion - Define and Export Module Variables
#=================================================================

#=================================================================
#region - Define PUBLIC Advanced Functions
#=================================================================
function Get-WSMMSIPropertyTable{
    <#
    .Synopsis
       Return the content of an MSI's "Property" table
    .EXAMPLE
        Get-MSIPropertyTable -Path 'C:\Application_1.0.msi'

        Name                           Value
        ----                           -----
        UpgradeCode                    {198663C2-2045-4B4E-B1D6-05414B4664CA}
        ALLUSERS                       1
        DISABLEADVTSHORTCUTS           1
        ARPPRODUCTICON                 Product.ico
        ARPCOMMENTS                    Application 1.0
        ARPCONTACT                     Vendor
        Manufacturer                   Vendor Corp.
        ProductCode                    {D5EDC5B4-1167-4368-970A-304328B4C80C}
        ProductLanguage                1033
        ProductName                    Application
        ProductVersion                 1.0.0.0
        SecureCustomProperties         UPGRADE;DOWNGRADE;CUSTOMPATH
        MsiHiddenProperties            RollBack;RollForward
    #>

    [CmdletBinding()]
    [OutputType([object])]
    Param
    (
        # Full path to the MSI file
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [validatescript({
            if (Get-ChildItem $_ | Where-Object Extension -eq '.MSI'){$true}
            else {Write-Error -Message "Invalid MSI Path / File Supplied: $_." -ErrorAction Stop}
        })]
        [string]$Path
    )

    Begin {
        # Unlike Get-Item/ChildItem, InvokeMember won't accept a relative path, so we'll convert if necessary
        $Path = Convert-Path -Path $Path

        # Attempt to connect to the Windows Installer COM Object
        try {
            $COMWindowsInstaller = New-Object -ComObject WindowsInstaller.Installer -ErrorAction Stop
        }
        catch {
            Write-Error $_ -ErrorAction Stop
        }

        # Define the DB query and binding flags
        $MSIDBQuery = 'SELECT * FROM Property'
        $InvokeMethod = [System.Reflection.BindingFlags]::InvokeMethod
        $GetProoperty = [System.Reflection.BindingFlags]::GetProperty
    }

    Process {
        # Open the MSI Database
        $MSIDatabase = $COMWindowsInstaller.GetType().InvokeMember(
            'OpenDatabase', 
            $InvokeMethod, 
            $Null, 
            $COMWindowsInstaller, 
            @($Path, 0)
        )

        # Load the query
        $View = $MSIDatabase.GetType().InvokeMember(
            'OpenView', 
            $InvokeMethod, 
            $null, 
            $MSIDatabase, 
            ($MSIDBQuery)
        )

        # Execute the query
        $View.GetType().InvokeMember(
            'Execute', 
            $InvokeMethod, 
            $null, 
            $View, 
            $null
        )

        # Iterate the return data and send output
        while ($Record = $View.GetType().InvokeMember('Fetch', $InvokeMethod, $null, $View, $null)) {
            $private:Name  = $Record.GetType().InvokeMember('StringData', $GetProoperty, $null, $Record, 1) 
            $private:Value = $Record.GetType().InvokeMember('StringData', $GetProoperty, $null, $Record, 2)
            @{$Name = $Value}
        }
    }
    End {
        # Clean up connections
        Remove-Variable COMWindowsInstaller,MSIDatabase,View,Record -Force -ErrorAction SilentlyContinue
    }
}
#=================================================================
#endregion - Define PUBLIC Advanced Functions
#=================================================================

#=================================================================
#region - Define PRIVATE Advanced Functions
#=================================================================
#=================================================================
#endregion - Define PRIVATE Advanced Functions
#=================================================================

#=================================================================
#region - Export Modules
#=================================================================
Export-ModuleMember -Function *-WSM*
#=================================================================
#endregion - Export Modules
#=================================================================