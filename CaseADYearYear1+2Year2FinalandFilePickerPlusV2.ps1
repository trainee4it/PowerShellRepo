function File-Picker
{

    #changed something else

    Add-Type -AssemblyName system.windows.forms

    $File = New-Object System.Windows.Forms.OpenFileDialog

    $File.InitialDirectory = "C:\files"

    $File.Filter = "csv files (*.csv)|*.csv|All files (*.*)|*.csv"

    $file.Title = "Please pick a csv file"

    $result = $File.ShowDialog()
    #added throw
    if($result -eq 'cancel')
    {

        throw "You have to pick a CSV file !"


    }

    



     $File.FileName


    



}




function Check-OU
{

    param($SamAccountName,$OU)

    $CorrectOU = $false

    $UserObject = Get-ADUser -Identity $SamAccountName

    if($UserObject.DistinguishedName.Split(',')[1] -eq ('OU='+$OU))
    {


        $CorrectOU = $true



    }

    $CorrectOU

}

function Test-AdUser
{
    
    param($Name)
    $Exist = $true
    try
    {

        $New = Get-ADUser -Identity $Name  -ErrorAction Stop 

    }

    catch
    {

        $Exist = $false

    }

    Write-Output -InputObject $Exist
}


function Test-AdOu
{
    
    param($Name)
    $Exist = $true

    $TestOU = 'OU='+$Name+',DC=contoso,dc=com'
    try
    {

        $New = Get-ADOrganizationalUnit -Identity $TestOU -ErrorAction Stop 

    }

    catch
    {

        $Exist = $false

    }

    Write-Output -InputObject $Exist
}

$UsersToImport = $null
$UsersToImport = Import-Csv -Path (File-Picker)

#Ophalen distinguishedname dus in dit geval dc=adatum,dc=com, bij jullie dc=contoso,dc=com
$RootDomain = (Get-ADDomain).distinguishedname

#ophalen suffix in dit geval adataum.com, bij jullie contoso.com
$UPNSUFFIX = (Get-ADDomain).forest

foreach($Item in $UsersToImport)
{

    #Krijg op het volgende commando foutmeldingen omdat
    #de OU uiteindelijk heel veel keer aangemaakt probeert te worden (77 keer)
    #Geen probleem, dit lossen we later op

    # +++++++++++++> Incorporeer hier de OU null check <++++++++++
    
    if((Test-AdOu -Name $Item.Description) -eq $true)
    {

        'Exists ' +$Item.Description


    }

    else
    {

        New-ADOrganizationalUnit -Name $Item.Description -ProtectedFromAccidentalDeletion $false

    }

    



    #Dit wordt de inlognaam zowel samaccountname als UPN
    $SamAccountName = $Item.FIRSTNAME[0] + $Item.LASTNAME
    
    $OU = 'OU='+$item.DESCRIPTION +','+$RootDomain 
    $UPN = $SamAccountName +'@'+$UPNSUFFIX
    $SecurePassword = ConvertTo-SecureString -String 'Pa55w.rd1234' -AsPlainText -Force
    $DisplayName = $Item.FIRSTNAME + ' '+ $Item.LASTNAME

    if((Test-AdUser -Name $SamAccountName) -eq $true)
    {

        Write-Verbose -Message ( 'Exists ' +$SamAccountName ) -Verbose

        if(Check-OU -SamAccountName $SamAccountName -OU $Item.DESCRIPTION)
        {

            Write-Verbose -Message ('User '+ $SamAccountName + ' is already in the correct OU ') -Verbose
 
        }
        else
        {
             Write-Verbose -Message ('User '+ $SamAccountName + ' is NOT in the correct OU ') -Verbose
            
            Get-ADUser -Identity $SamAccountName | Move-ADObject -TargetPath $OU

        }
    }

    else
    {

        New-ADUser -Name $DisplayName -UserPrincipalName $UPN -SamAccountName $SamAccountName -Enabled $true -AccountPassword $SecurePassword -GivenName $Item.FIRSTNAME -Surname $Item.LASTNAME -Path $OU
    
        Write-Verbose -Message ("creating user " + $SamAccountName) -Verbose

    }

   


}