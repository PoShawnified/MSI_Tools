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
function Get-MSIPropertyTable{
    <#
    .Synopsis
       Return the content of an MSI's "Property" table
    .EXAMPLE
        Get-MSIPropertyTable -Path 'C:\Application_1.0.msi' -Transforms ('C:\AppMst1.mst','C:\AppMst2.mst')

        Property                       Value                                       Public
        --------                       -----                                       ------
        UpgradeCode                    {198663C2-2045-4B4E-B1D6-05414B4664CA}      False
        ALLUSERS                       1                                           True
        DISABLEADVTSHORTCUTS           1                                           True
        ARPPRODUCTICON                 Product.ico                                 True
        ARPCOMMENTS                    Application 1.0                             True
        ARPCONTACT                     Vendor                                      True
        Manufacturer                   Vendor Corp.                                False
        ProductCode                    {D5EDC5B4-1167-4368-970A-304328B4C80C}      False
        ProductLanguage                1033                                        False
        ProductName                    Application                                 False
        ProductVersion                 1.0.0.0                                     False
        SecureCustomProperties         UPGRADE;DOWNGRADE;CUSTOMPATH                False
        MsiHiddenProperties            RollBack;RollForward                        False
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
            if (Get-Item $_ | Where-Object Extension -eq '.MSI'){$true}
            else {Write-Error -Message "Invalid MSI Path / File Supplied: $_." -ErrorAction Stop}
        })]
        [string]$Path,

        # Full path to the MST file
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [validatescript({
            $_ | Foreach-Object { 
              $thisMst = $_
              if (Get-Item $thisMst | Where-Object Extension -eq '.MST'){$true}
              else {Write-Error -Message "Invalid Transform Path / File Supplied: $thisMst." -ErrorAction Stop}
            }
        })]
        [string[]]$Transforms
    )

    Begin {
        # Unlike Get-Item/ChildItem, InvokeMember won't accept a relative path, so we'll convert if necessary
        $Path = (Convert-Path -Path $Path).ToString()

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
        $GetProperty = [System.Reflection.BindingFlags]::GetProperty
    }

    Process {
        # Open the MSI Database
        $MSIDatabase = $COMWindowsInstaller.GetType().InvokeMember(
            'OpenDatabase', 
            $InvokeMethod, 
            $null, 
            $COMWindowsInstaller, 
            @($Path, 0)
        )
 
        # Apply MSTs
        if ($Transforms){
            $Transforms | ForEach-Object {
                # Unlike Get-Item/ChildItem, InvokeMember won't accept a relative path, so we'll convert if necessary
                $thisMST = (Convert-Path -Path $_).ToString()

                $MSIDatabase.GetType().InvokeMember(
                    'ApplyTransform',
                    $InvokeMethod,
                    $null,
                    $MSIDatabase,
                    @($thisMST, 0)
                )
            }
        }
        
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
            $private:Property  = $Record.GetType().InvokeMember('StringData', $GetProperty, $null, $Record, 1) 
            $private:Value = $Record.GetType().InvokeMember('StringData', $GetProperty, $null, $Record, 2)

            # Public properties are always upper-case
            [PSCustomObject]@{
                Property = $Property
                Value = $Value
                Public = if ($Property -ceq $Property.ToUpper()){$true}else{$false}
            }
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
Export-ModuleMember -Function *-*
#=================================================================
#endregion - Export Modules
#=================================================================