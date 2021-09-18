<#
.SYNOPSIS
		<This script generate multiple configuration files for Avaya VoIP phones>
.INPUTS
		<An xlsx file>
.NOTES
		Version:        1.0
		Author:         Simon Baleine
		Creation Date:  Tuesday, December 1, 2020
		Purpose/Change: Initial script development

.EXAMPLE
		< ./main.ps1 ./phones.csv>
#>

# To-Do
# Handle people with the same names

#-----------------------------------------------------------[Functions]------------------------------------------------------------

$EXCEL_FILE_LOCATION = ".\employes.xlsx";
$DOMAIN_NAME = "NEVADA";
$TOP_LEVEL_DOMAIN_NAME = "US";

$SELECTED_WORKSHEET_NUMBER = $args[0];

function Generate-Username {
	param ([string]$firstName, [string]$lastName)
	$firstName = [Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding(1251).GetBytes($firstName)).Trim().ToLower().Replace(' ','');
	$lastName = [Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding(1251).GetBytes($lastName)).Trim().ToLower().Replace(' ','');
	$username = -join($firstName.ToCharArray()[0], ".", $lastName);

	return $username;
}

function Get-RandomPassword {
    Param(
        [Parameter(mandatory = $true)]
        [int]$Length
    )
    Begin {
        if ($Length -lt 4) {
            End
        }
        $Numbers = 1..9
        $LettersLower = 'abcdefghijklmnopqrstuvwxyz'.ToCharArray()
        $LettersUpper = 'ABCEDEFHIJKLMNOPQRSTUVWXYZ'.ToCharArray()
        $Special = '!@#$%^&*()=+[{}]/?<>'.ToCharArray()

        # For the 4 character types (upper, lower, numerical, and special)
        $N_Count = [math]::Round($Length * .2)
        $L_Count = [math]::Round($Length * .4)
        $U_Count = [math]::Round($Length * .2)
        $S_Count = [math]::Round($Length * .2)
    }
    Process {
        $Pswrd = $LettersLower | Get-Random -Count $L_Count
        $Pswrd += $Numbers | Get-Random -Count $N_Count
        $Pswrd += $LettersUpper | Get-Random -Count $U_Count
        $Pswrd += $Special | Get-Random -Count $S_Count

        # If the password length isn't long enough (due to rounding), add X special characters
        # Where X is the difference between the desired length and the current length.
        if ($Pswrd.length -lt $Length) {
            $Pswrd += $Special | Get-Random -Count ($Length - $Pswrd.length)
        }

        # Lastly, grab the $Pswrd string and randomize the order
        $Pswrd = ($Pswrd | Get-Random -Count $Length) -join ""
    }
    End {
        $Pswrd
    }
}

function Get-WorkSheet($workSheetNumber, $test) {
	$workSheetName = "liste" + $workSheetNumber;
	return (Import-Excel -Path $EXCEL_FILE_LOCATION -WorkSheetname $workSheetName);
}

function Get-UsersFromWorkSheet {
	param ($worksheet)
	$users = New-Object System.Collections.Generic.List[System.Object];
	$usernames = New-Object System.Collections.Generic.List[System.Object];
	$usersDuplicates = New-Object System.Collections.Generic.List[System.Object];

	foreach ($u in $worksheet) {
		$user = New-Object -TypeName PSCustomObject;
		$user | Add-Member -MemberType NoteProperty -Name "FirstName" -Value $u.'Prénom';
		$user | Add-Member -MemberType NoteProperty -Name "LastName" -Value $u.'Nom';
		$user | Add-Member -MemberType NoteProperty -Name "Description" -Value $u.'Description';
		$user | Add-Member -MemberType NoteProperty -Name "InternalNumber" -Value $([int]($u.'N° Interne'));
		$user | Add-Member -MemberType NoteProperty -Name "Desk" -Value $u.'Bureau';
		$user | Add-Member -MemberType NoteProperty -Name "OrganizationalUnit" -Value $u.'Département'.Split('/');
		$userPath = [string]::Format("OU={0},OU={1}",$user.OrganizationalUnit[0], $user.OrganizationalUnit[1]);
		$user | Add-Member -MemberType NoteProperty -Name "Path" -Value $userPath;
		$username =  $(Generate-Username -firstName $u.'Prénom' -lastName $u.'Nom');
		$user | Add-Member -MemberType NoteProperty -Name "Username" -Value $username;
		$user | Add-Member -MemberType NoteProperty -Name "IsManager" -Value $u.'Responsable';

		$passwordLength = 7;
		# Setup a longer passwordLength if the user belong in the Direction OU
		if ($user.OrganizationalUnit[0] -like "Direction") {
			$passwordLength = 15;
		}
		$user | Add-Member -MemberType NoteProperty -Name "Password" -Value $(Get-RandomPassword -Length $passwordLength);

		if ($username -Notin $usernames) {
			$users.Add($user);
			$usernames.Add($username);
		}
		else {
			$usersDuplicates.Add($user);
		}

	}
	return @($users,$usersDuplicates);
}

function Get-OrganizationalUnitsPaths {
	param ($worksheet)
	# Select every cells belonging in the 'Département' columns accross every worksheets
	$rawDepartements = ($worksheet | Select-Object -Unique 'Département');

	# Create a list that will hold every child OU names
	$childOrganizationalsUnitsFullPaths = New-Object System.Collections.Generic.List[System.Object];

	foreach ($rawDepartement in $rawDepartements) {
		# Get the root OU name by splitting the string on the "/" character and taking the last element
		$split = $rawDepartement.'Département'.Split('/');

		if ($rawDepartement.'Département' -match "/") {
			$childOrganizationalsUnitsFullPaths.Add(@($split[-1], $split[0]));
		}

		$childOrganizationalsUnitsFullPaths.Add($split[-1]);
	}

	# Convert the list to an array
	$childOrganizationalsUnitsFullPaths = $childOrganizationalsUnitsFullPaths.ToArray();

	# Remove duplicates from that list
	$childOrganizationalsUnitsFullPaths = ($childOrganizationalsUnitsFullPaths | Select-Object -Unique);

	return $childOrganizationalsUnitsFullPaths;
}

function Get-OrganizationalUnitNames {
	param ($worksheet)
	$organizationalUnitNames = New-Object System.Collections.Generic.List[System.Object];
	Get-OrganizationalUnitsPaths -worksheet $worksheet | ForEach-Object -Process { $organizationalUnitNames += $_ }

	return ($organizationalUnitNames | Select-Object -Unique)
}

function Get-OrganizationalUnits {
	param ($worksheet)
	$organizationalUnitsPaths = Get-OrganizationalUnitsPaths -worksheet $worksheet;
	$organizationalUnits = New-Object System.Collections.Generic.List[System.Object];

	foreach ($organizationalUnitPath in $organizationalUnitsPaths) {
		$organizationalUnit = New-Object -TypeName PSCustomObject;
		
		# Result is in the form of array [OU, OU]
		if ($organizationalUnitPath -is [System.Array]) {
			$organizationalUnit | Add-Member -MemberType NoteProperty -Name "Name" -Value $organizationalUnitPath[1];
			$organizationalUnit | Add-Member -MemberType NoteProperty -Name "Location" -Value $organizationalUnitPath;
		}
		# Result is in the form of string OU
		else {
			$organizationalUnit | Add-Member -MemberType NoteProperty -Name "Name" -Value $organizationalUnitPath;
			$organizationalUnit | Add-Member -MemberType NoteProperty -Name "Location" -Value $organizationalUnitPath;
		}
		$organizationalUnits.Add($organizationalUnit);
	}

	return $organizationalUnits
}

function Get-GlobalGroups {
	param ($worksheet)

	$globalGroups = New-Object System.Collections.Generic.List[System.Object];
	$organizationalUnitNames = Get-OrganizationalUnitNames -worksheet $worksheet;
	
	foreach ($organizationalUnitName in $organizationalUnitNames) {
		$globalGroup = New-Object -TypeName PSCustomObject;

		$groupName = $organizationalUnitName.Replace(' ','');
		$groupName = "GG_$groupName";
		$globalGroup | Add-Member -MemberType NoteProperty -Name "Name" -Value $groupName;
		$globalGroup | Add-Member -MemberType NoteProperty -Name "Location" -Value "Groupes";
		$globalGroup | Add-Member -MemberType NoteProperty -Name "GroupScope" -Value "Global";
		$globalGroups.Add($globalGroup);
		
		$respGlobalGroup = New-Object -TypeName PSCustomObject;
		$respGroupName = $organizationalUnitName.Replace(' ','');
		$respGroupName = "GG_Resp$respGroupName";
		$respGlobalGroup | Add-Member -MemberType NoteProperty -Name "Name" -Value $respGroupName;
		$respGlobalGroup | Add-Member -MemberType NoteProperty -Name "Location" -Value "Groupes";
		$respGlobalGroup | Add-Member -MemberType NoteProperty -Name "GroupScope" -Value "Global";
		$globalGroups.Add($respGlobalGroup);
	}

	return ($globalGroups);
}

function Get-LocalGroups {
	param ($worksheet)
	
	$localGroups = New-Object System.Collections.Generic.List[System.Object];
	$globalGroups = Get-GlobalGroups -worksheet $worksheet;
	foreach ($globalGroup in $globalGroups) {
		$localReadGroup = New-Object -TypeName PSCustomObject;
		$localReadGroup | Add-Member -MemberType NoteProperty -Name "Name" -Value $(($globalGroup.Name).Replace('GG_','GL_R_partage'));
		$localReadGroup | Add-Member -MemberType NoteProperty -Name "Location" -Value "Groupes";
		$localReadGroup | Add-Member -MemberType NoteProperty -Name "GroupScope" -Value "DomainLocal";

		$localReadWriteGroup = New-Object -TypeName PSCustomObject;
		$localReadWriteGroup | Add-Member -MemberType NoteProperty -Name "Name" -Value $(($globalGroup.Name).Replace('GG_','GL_RW_partage'));
		$localReadWriteGroup | Add-Member -MemberType NoteProperty -Name "Location" -Value "Groupes";
		$localReadWriteGroup | Add-Member -MemberType NoteProperty -Name "GroupScope" -Value "DomainLocal";

		$localGroups.Add($localReadWriteGroup);
		$localGroups.Add($localReadGroup);
	}
	return $localGroups
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------



# Get-OrganizationalUnits -worksheet $worksheet;
for ($i = 4; $i -lt 8; $i++) {
	$worksheet = Get-WorkSheet($i);
	
	$results = Get-UsersFromWorkSheet -worksheet $worksheet;
	$duplicates = $results[1]
	$users = $results[0]
	
	# Retrieve users with a username that's too big
	$tooBigUsernames = ($users | Where-Object { $_.Username.Length -gt 20 }); 
	
	# Remove users with a username that's too big fron the $users array
	$users = $users | Where-Object {$_ -notin $tooBigUsernames}
	
	$users | Export-Clixml ./output/$i-users.xml;
	Get-OrganizationalUnits -worksheet $worksheet | Export-Clixml ./output/$i-ous.xml;
	$tooBigUsernames | ConvertTo-Json > ./output/$i-too-long-usernames.json;
	$duplicates | ConvertTo-Json > ./output/$i-duplicate-usernames.json;
	$(Get-GlobalGroups -worksheet $worksheet) + $(Get-LocalGroups -worksheet $worksheet) | Export-Clixml ./output/$i-groups.xml;
}
