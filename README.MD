# Lancer sur WYOMING
Import-Module -Force .\ImportUsers\ImportUsers.psm1;
Import-ADConfigFromXML -Verbose -GroupXMLFile .\output\6-groups.xml -UserXMLFile .\output\4-users.xml -OrganizationalUnitXMLFile .\output\4-ous.xml -RootOU WYOMING -Path "DC=WYOMING,DC=NEVADA,DC=NET";

# Lancer sur WASHINGTON
Import-Module -Force .\ImportUsers\ImportUsers.psm1;
Import-ADConfigFromXML -Verbose -GroupXMLFile .\output\6-groups.xml -UserXMLFile .\output\5-users.xml -OrganizationalUnitXMLFile .\output\5-ous.xml -RootOU WASHINGTON -Path "DC=WASHINGTON,DC=NEVADA,DC=NET";

# Lancer sur NEVADA
Import-Module -Force .\ImportUsers\ImportUsers.psm1;
Import-ADConfigFromXML -Verbose -GroupXMLFile .\output\6-groups.xml -UserXMLFile .\output\6-users.xml -OrganizationalUnitXMLFile .\output\6-ous.xml -RootOU NEVADA -Path "DC=NEVADA,DC=NET";
