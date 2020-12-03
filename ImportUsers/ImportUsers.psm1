<#
 .Synopsis
  Add calendar.

 .Description
  Displays a visual representation of a calendar. This function supports multiple months
  and lets you highlight specific date ranges or days.

 .Parameter Start
  The first month to display.

 .Parameter End
  The last month to display.

 .Parameter FirstDayOfWeek
  The day of the month on which the week begins.

 .Parameter HighlightDay
  Specific days (numbered) to highlight. Used for date ranges like (25..31).
  Date ranges are specified by the Windows PowerShell range syntax. These dates are
  enclosed in square brackets.

 .Parameter HighlightDate
  Specific days (named) to highlight. These dates are surrounded by asterisks.

 .Example
   # Show a default display of this month.
   Show-Calendar

 .Example
   # Display a date range.
   Show-Calendar -Start "March, 2010" -End "May, 2010"

 .Example
   # Highlight a range of days.
   Show-Calendar -HighlightDay (1..10 + 22) -HighlightDate "December 25, 2008"
#>

Import-Module ActiveDirectory;

function New-ADOrganizationalUnitIfNotExist {
	param([string]$name, [string]$Path, [string]$description, [String]$ProtectedFromAccidentalDeletion = $false)
    $ouDN = "OU=$name,$Path";

    try {
        Get-ADOrganizationalUnit -Identity $ouDN | Out-Null;
        Write-Verbose "OU '$ouDN' already exists.";
    }
    catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
        Write-Verbose "Creating new OU '$ouDN'";
        New-ADOrganizationalUnit -Name $name -Path $Path -Description $description;
    }
}

function New-ADGroupIfNotExist {
	param([string]$name, [string]$Path, [string]$description, [string]$GroupScope)
    $groupDN = "CN=$name,$Path";

    try {
        Get-ADGroup -Identity $groupDN | Out-Null;
        Write-Verbose "Group $groupDN already exists";
    }
    catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
        Write-Verbose "Creating new Group '$groupDN'";
        New-ADGroup -Name $name -Path $Path -Description $description -GroupScope $GroupScope;
    }
}
function New-ADUserIfNotExist {
	param(
        [string]$Name,
        [string]$GivenName,
        [string]$Surname,
        [string]$SamAccountName,
        [string]$Path,
        [string]$UserDNPath,
        $OfficePhone,
        [securestring]$AccountPassword = ("Test1234=" | ConvertTo-SecureString -AsPlainText -Force)
        )
    try {
        Get-ADUser -Identity $UserDNPath | Out-Null;

        Write-Verbose "User $SamAccountName already exists";
    }
    catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
        Write-Verbose "Creating new User $UserDNPath";
        New-ADUser -Name $SamAccountName `
        -GivenName $GivenName `
        -Surname $Surname `
        -SamAccountName $SamAccountName `
        -Path $Path `
        -Enabled $true `
        -AccountPassword $AccountPassword `
        -OfficePhone $OfficePhone
    }
}

function Import-ADConfigFromXML {
	param(
		[Parameter(Mandatory=$true)]
		[String]$GroupXMLFile,
		[Parameter(Mandatory=$true)]
		[String]$UserXMLFile,
		[Parameter(Mandatory=$true)]
		[String]$organizationalUnitXMLFile,
		[Parameter(Mandatory=$true)]
		[String]$RootOU,
		[Parameter(Mandatory=$true)]
		[String]$Path,
		[String]$ProtectedFromAccidentalDeletion = $false
		)

	$groups = Import-Clixml -Path $GroupXMLFile;    
	$users = Import-Clixml -Path $UserXMLFile;    
	$organizationalUnits = Import-Clixml -Path $OrganizationalUnitXMLFile;
	$pathParsed = $Path.Replace('DC=','').Split(',');
	$rootOUIdentity = [string]::Format("OU={0},{1}",$RootOU, $Path);
	$groupesOUIdentity = [string]::Format("OU=Groupes,OU={0},{1}",$RootOU, $Path);
	$groupesOUpath = [string]::Format("OU={0},{1}", $RootOU, $Path);

	# Create the root Organizational Unit if it does not exist
    New-ADOrganizationalUnitIfNotExist -ProtectedFromAccidentalDeletion $ProtectedFromAccidentalDeletion -Name $RootOU -Path $Path;

    # "OU=Groupes,OU=Whalewave,DC=WHALEWAVE,DC=NET"
	New-ADOrganizationalUnitIfNotExist -ProtectedFromAccidentalDeletion $ProtectedFromAccidentalDeletion -Name "Groupes" -Path $groupesOUpath;

	# Create every groups
	foreach ($group in $groups) {
        New-ADGroupIfNotExist -Path $groupesOUIdentity -Name $($group.Name) -GroupScope $($group.GroupScope);
	}

	# Create every Root Organizational Units
	foreach ($organizationalUnit in $organizationalUnits | Where-Object Location -Is [string]) {
		$organizationalUnitPath = $rootOUIdentity;
		New-ADOrganizationalUnitIfNotExist -ProtectedFromAccidentalDeletion $ProtectedFromAccidentalDeletion -Name $organizationalUnit.Name -Path $organizationalUnitPath;
	}
	
	# Create every Child Organizational Units
	foreach ($organizationalUnit in $organizationalUnits | Where-Object Location -IsNot [string]) {
        $organizationalUnitPath = [string]::Format("OU={0},{1}",$organizationalUnit.Location[0], $rootOUIdentity);
		New-ADOrganizationalUnitIfNotExist -ProtectedFromAccidentalDeletion $ProtectedFromAccidentalDeletion -Name $organizationalUnit.Name -Path $organizationalUnitPath;
    }
    
    foreach ($user in $users) {
        # Format paths for user creation 
        $userOUPath = [string]::Format("{0},{1}", $user.Path, $rootOUIdentity).Replace('OU=,','');
        $userDNPath = [string]::Format("CN={0},{1}", $user.Username,$userOUPath).Replace('OU=,','');

        if ($user.OrganizationalUnit -Is [string]) {
            $userGroupName = [string]::Format("GG_{0}", $user.OrganizationalUnit);
        }
        elseif ($user.IsManager) {
            $userGroupName = [string]::Format("GG_Resp{0}", $user.OrganizationalUnit[0]);
        } 
        else {
            $userGroupName = [string]::Format("GG_{0}", $user.OrganizationalUnit[0]);
        }

        $userName = [string]::Format("{0} {1}", $user.FirstName,$user.Lastname); 
        
        # Add the new user
        New-ADUserIfNotExist -Name $userName `
        -UserDNPath $userDNPath `
        -GivenName $user.FirstName `
        -Surname $user.LastName `
        -SamAccountName $user.Username `
        -Path $userOUPath `
        -AccountPassword ($user.Password | ConvertTo-SecureString -AsPlainText -Force)`
        -OfficePhone $user.InternalNumber
        
        # Add user to its group
        $adUser = Get-ADUser -Identity $userDNPath;
        $adGroup = Get-ADGroup -Filter { Name -Like $userGroupName }
        $adGroup | Add-ADGroupMember -Members $adUser;
    }

}
Export-ModuleMember -Function Import-ADConfigFromXML
