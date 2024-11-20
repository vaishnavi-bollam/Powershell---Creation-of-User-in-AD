$id="672d968eafce399fefa7ec17"

$flag=0
$global:tempflag=0
$worknotes=$null

try{
    Import-Module -Name "IexecuteModules" -ErrorAction Stop
    $data=Get-IxMongoData -id $id
    $worknotes+= Set-IxWorkNotesText "Inputs Extracted  Successfully"
}catch{
    $_
    $worknotes+= Set-IxWorkNotesText -Flag ERROR -Note "Inputs Extraction Failed"
    $flag = '2'
}

$name                  = $data.validated_inputs.'parent.variables.name'
#$name="asssww Test"
$title                 = $data.validated_inputs.'parent.variables.title'
$organization          = $data.validated_inputs.'parent.variables.which_organization_are_the_user_to_work_for'
#$organization="VKR Holding A/S"
$location              = $data.validated_inputs.'parent.variables.location'
$manager               = $data.validated_inputs.'parent.variables.manager'
$department            = $data.validated_inputs.'parent.variables.departement'
$shortname             = $data.validated_inputs.'parent.variables.requested_shortname'
#$shortname="asw.vkr"
$ext_Email           = $data.validated_inputs.'parent.variables.email'
#$ext_Email="bollam.vaishnavi@hcltech.com"

$Ritm = $data.validated_inputs.'parent.number'


$sys_id                = $data.validated_inputs.sys_id
$table                 = "sc_task"
$assignment_group      = $data.fulfillment_group_sysid
$short_description = $data.short_description

$cred = Get-IxServiceAccount
$worknote=Set-IxWorkNotesText "Automation work in progess"
try{Update-IxTicket -sys_id $sys_id -table $table -work_notes $worknote -case "wip"}catch{$_}
try{Update-IxMongoData -id $id -status WorkInProgress}catch{$_}


function set-userproxy {

    param([string] $userLogonName1)

    $User_id = $userLogonName1

    $userLogonName = ($User_id -split '\.')[0]


    $mail = "$userLogonName$userLogonNamedomain"

    $mailnickname = $User_id

    $targetaddress = "$userLogonName@domino.velux.com"

    $SMTP = "SMTP:"

    $sip = "sip:"

    $smtp2 = "smtp:"

    $proxyaddresses = @(

        "$smtp2$userLogonName@velux.es",

        "$smtp2$userLogonName@dakea.com",

        "$smtp2$userLogonName@velux.nl",

        "$smtp2$userLogonName@velux.com.ar",

        "$smtp2$userLogonName@velux.com.au",

        "$smtp2$userLogonName@velux.no",

        "$smtp2$userLogonName@velux.ru",

        "$smtp2$userLogonName@velux.mail.onmicrosoft.com",

        "$smtp2$userLogonName@veluxfondene.dk",

        "$smtp2$userLogonName@velux.se",

        "$smtp2$userLogonName@velux.by",

        "$smtp2$userLogonName@velux.ca",

        "$smtp2$userLogonName@veluxfonden.dk",

        "$smtp2$userLogonName@balio.org",

        "$smtp2$userLogonName@velux.co.uk",

        "$smtp2$userLogonName@velux.ch",

        "$smtp2$userLogonName@rooflite.com",

        "$smtp2$userLogonName@itzala.com",

        "$smtp2$userLogonName@velux.cz",

        "$smtp2$userLogonName@velux.pl",

        "$smtp2$userLogonName@velux.de",

        "$smtp2$userLogonName@veluxshop.com",

        "$smtp2$userLogonName@velux.dk",

        "$smtp2$userLogonName@itzala.net",

        "$smtp2$userLogonName@vinduet.dk",

        "$smtp2$userLogonName@veluxfoundations.dk",

        "$smtp2$userLogonName@velux.co.nz",

        "$smtp2$userLogonName@velux.ie",

        "$smtp2$userLogonName@villumfonden.dk",

        "$smtp2$userLogonName@altaterra.eu"

    )
 

    $newProxyAddresses = @(

        "$sip$userLogonName$userLogonNamedomain",

        "$SMTP$userLogonName$userLogonNamedomain"

    )

    try {

        $currentUser = Get-ADUser -Identity $User_id -Properties proxyAddresses

        $existingProxyAddresses = $currentUser.proxyAddresses | Where-Object { $_ -notmatch $userLogonNamedomain }


        $finalProxyAddresses = @(

            $existingProxyAddresses + $proxyaddresses + $newProxyAddresses

        ) | ForEach-Object { [string]$_ }

        Set-ADUser -Identity $User_id -Replace @{

            mail = $mail;

            mailNickname = $mailnickname;

            targetAddress = $targetaddress;

            

        }


        $existingValues = Get-ADUser -Identity $User_id -Properties proxyAddresses | Select-Object -ExpandProperty proxyAddresses

        foreach ($address in $existingValues) {

            Set-ADUser -Identity $User_id -Remove @{proxyAddresses = $address}

        }

        Set-ADUser -Identity $User_id -Add @{proxyAddresses = $finalProxyAddresses}
 

        $setproxy = Set-IxWorkNotesText -Flag SUCCESS -Note "Successfully updated proxy_addresses for User : [$(Format-IxText $User_id)]"

    } catch 
    {

        Write-Host $_

        $setproxy = Set-IxWorkNotesText -Flag ERROR -Note "Failed to update the proxy address for User : [$(Format-IxText $User_id)]"

        $global:tempflag = 1

    }

    return $setproxy

}

function sendmaildetails
{
    param([string] $To_address)
    $from        = "AutomationTeam@velux.com"
    $to          = $To_address
    $subject     = "[$($ritm)] - AD User Account - Request Internal user account - $($name)"
    $SmtpServer  = "mailrelay.velux.org"
    $body        = "<html>  
                    <body style='color:#455A64;font-family: `"Calibri Light`";'>
                    <p>Hi <b>Requester</b>,</p>
                    <div>
                        <p>The requested AD internal user account has been created and will be available within the next 2 hours</p>
                    </div>
			        <ul style='list-style-type: none;font-size:14px;'>
			        <li>FULL NAME &emsp;&nbsp;&nbsp;&nbsp;: <b>$($name) </b></li>
                        <li>USER ID&emsp;&emsp;&emsp;&nbsp;&nbsp;: <b>$($userLogonName) </b></li>
                    </ul>
		            <p>Click <a href='https://velux.service-now.com/sp?id=sc_cat_item_guide&sys_id=9f758b5d1bf628103d3510a1b24bcbb2' style='color:blue;'><b>here</b></a> to request basic access such as Office 365, Citrix Web Access or hardware.</p>
		
		            <p>Click <a href='https://velux.service-now.com/sp' style='color:blue;'><b>here</b></a> to request more specific access not available in the basic form raise a separate request via the Service Portal</p> 
                    <br>
                    <em style='color:#1565c0;font-size:11px;'>***Please note this is auto-generated mail. Do not Reply to this mail***</em>
                    <br><br><br>
                    <em style='font-size:12px;'>Thanks & Regards,</em>
                    <p style='color:#1565c0;'><b>VELUX AUTOMATION TEAM</b></p>
		
                    </body>
                    </html>"
    try
    {
        Send-MailMessage -From $from -to $to -Subject $subject -SmtpServer $SmtpServer -Body $body -BodyAsHtml -Port 587 -UseSsl
        $commentmail= Set-IxWorkNotesText -Flag SUCCESS -Note "Account details sent to [$(Format-IxText $To_address)]"
    }
    catch
    {
        Write-Host $_
        $commentmail= Set-IxWorkNotesText -Flag ERROR -Note "Unable to send the Password"
        $global:tempflag=1
    }

    return $commentmail
}

function sendmailpassword
{
    param([string] $To_address)
    $from        = "AutomationTeam@velux.com"
    $to          = $To_address
    $subject     = "[$($ritm)] - AD User Account - Request internal user account - $($name)"
    $SmtpServer  = "mailrelay.velux.org"
    $body        = "<html>  
                    <body style='color:#455A64;font-family: `"Calibri Light`";'>
                    <p>Hi <b>Requester</b>,</p>
                    <div>
                        <p>As per the request (<b>$($ritm)</b>) we have successfully  created the User account for <b>$($name)</b> will be available <b style='color:#f57f17;'>within/after 2 hours</b></p>
                    </div>
                    <ul style='list-style-type: none;font-size:14px;'>
                        <li>USER ID&emsp;&nbsp;&nbsp;&nbsp;: <b>$($userLogonName) </b></li>
                        <li>PASSWORD&nbsp;: <b>$($password)</b></li>
                    </ul>
                    <p>Please note the password is case sensitive and there is no space before and after the Characters</p>
                    <em style='color:#2979FF;font-size: 13px;'>Note : This password is temporary password and will be valid only for next login from now. Your New password should be at least 10 alphanumeric characters long & Contain both upper and lower-case letters(e.g., a-z, A-Z).Do not contain words, personal information, company related words or other reconcilable combinations characters (e.g. slang, company department).</em>
                    <br>
                    <em style='color:#1565c0;font-size:11px;'>***Please note this is auto-generated mail. Do not Reply to this mail***</em>
                    <br><br><br>
                    <em style='font-size:12px;'>Thanks & Regards,</em>
                    <p style='color:#1565c0;'><b>VELUX AUTOMATION TEAM</b></p>
                    </body>
                    </html>"
    try
    {
        Send-MailMessage -From $from -to $to -Subject $subject -SmtpServer $SmtpServer -Body $body -BodyAsHtml -Port 587 -UseSsl
        $commentmail= Set-IxWorkNotesText -Flag SUCCESS -Note "Password sent to [$(Format-IxText $To_address)]"
    }
    catch
    {
        Write-Host $_
        $commentmail= Set-IxWorkNotesText -Flag ERROR -Note "Unable to send the Password"
        $global:tempflag=1
    }

    return $commentmail
}


if($short_description -like "*Create O365 account*"){

$Password        = $(Get-IxRandomPassword -length 12)

if ($name -match "\s") {
    $nameParts = $name -split "\s", 2
    $firstname = $nameParts[0]
    $lastname = $nameParts[1]
} else {
    $firstname = $name
    $lastname = ""
}

$domainName=""
#$userLogonName=$shortname
$fullName = "$firstName $lastName"

$OUpath = ""

if ($Organization -eq "VKR Holding A/S") {
    $OUpath = 'OU=Standard Users,OU=Users,OU=DNKHR,OU=Locations,DC=velux,DC=org'
} else {
    $OUpath = 'OU=Standard Users,OU=Users,OU=DNKSR,OU=Locations,DC=velux,DC=org'
}

#$email="$firstName@vkr-holding.com"



$userLogonNamedomain=""

switch ($Organization) {
    "VKR Holding A/S" {
        $domainName = ".vkr"
        $userLogonNamedomain = "@vkr-holding.com"
    }
    "Villum Fonden" {
        $domainName = ".fon"
        $userLogonNamedomain = "@villumfonden.dk"
    }
    "VELUX Fonden" {
        $domainName = ".fon"
        $userLogonNamedomain = "@veluxfonden.dk"
    }
    "ADDEK" {
        $domainName = ".fond"
        $userLogonNamedomain = "@velux.com"
    }
    "Fondene" {
        $domainName = ".fon"
        $userLogonNamedomain = "@fondene.dk"
    }
    default {
        Write-Host "Unknown organization"
    }
}

function Generate-InternalUserID{
    Param($name)
    $name = $name -replace '[^a-zA-Z]', ''
    foreach($index in $(0..$($name.Length-1))){
        #$uid=
        $uid = $name.Substring($index,3)
        $user1 = $uid[0]+$uid[1]+$uid[2]+$domainName
        $user2 = $uid[0]+$uid[2]+$uid[1]+$domainName
        $user3 = $uid[1]+$uid[0]+$uid[2]+$domainName
        $user4 = $uid[1]+$uid[2]+$uid[0]+$domainName
        $user5 = $uid[2]+$uid[0]+$uid[1]+$domainName
        $user6 = $uid[2]+$uid[1]+$uid[0]+$domainName
        Write-Host "Tried:($uid) $user1,$user2,$user3,$user4,$user5,$user6" -ForegroundColor DarkYellow
        try{ Get-ADUser -Identity $user1 -ErrorAction SilentlyContinue|out-null;Write-Host "$User1 Exists in AD" -ForegroundColor cyan}catch{if($(Get-IxExistingSnowSysUser -searchvalue $user1)){return $user1;continue}}
        try{ Get-ADUser -Identity $user2 -ErrorAction SilentlyContinue|out-null;Write-Host "$User2 Exists in AD" -ForegroundColor cyan}catch{if($(Get-IxExistingSnowSysUser -searchvalue $user2)){return $user2;continue}}
        try{ Get-ADUser -Identity $user3 -ErrorAction SilentlyContinue|Out-Null;Write-Host "$User3 Exists in AD" -ForegroundColor cyan}catch{if($(Get-IxExistingSnowSysUser -searchvalue $user3)){return $user3;continue}}
        try{ Get-ADUser -Identity $user4 -ErrorAction SilentlyContinue|Out-Null;Write-Host "$User4 Exists in AD" -ForegroundColor cyan}catch{if($(Get-IxExistingSnowSysUser -searchvalue $user4)){return $user4;continue}}
        try{ Get-ADUser -Identity $user5 -ErrorAction SilentlyContinue|Out-Null;Write-Host "$User5 Exists in AD" -ForegroundColor cyan}catch{if($(Get-IxExistingSnowSysUser -searchvalue $user5)){return $user5;continue}}
        try{ Get-ADUser -Identity $user6 -ErrorAction SilentlyContinue|Out-Null;Write-Host "$User6 Exists in AD" -ForegroundColor cyan}catch{if($(Get-IxExistingSnowSysUser -searchvalue $user6)){return $user6;continue}}
    }
}

$userLogonName1 = $shortname
#$userLogonName1="amb.vkr"

try {
    if (Get-ADUser -Identity $userLogonName1 -ErrorAction SilentlyContinue) {
        Write-Host "$userLogonName1 Exists in AD" -ForegroundColor cyan
        $userLogonName1=Generate-InternalUserID -name $fullName
    } else {
        if (Get-IxExistingSnowSysUser -searchvalue $userLogonName1) {
            Write-Host "$userLogonName1 Exists in Snow" -ForegroundColor cyan
            $userLogonName1=Generate-InternalUserID -name $fullName
        }
    }
} catch {
    #Write-Host return $user1;continue
}


#$userLogonName1 = Generate-InternalUserID -name $fullName
$userLogonName = ($userLogonName1 -split '\.')[0] 




$folderPath = "C:\temp\$Ritm"
$textFilePath="$folderPath\$Ritm.txt"
 

if (!(Test-Path -Path $folderPath)) {

    New-Item -ItemType Directory -Path $folderPath
    New-Item -ItemType File -Path $textFilePath -Force | Out-Null 
    @($userLogonName1, $userLogonName, $userLogonNamedomain) | Set-Content -Path $textFilePath
    Write-Output "Text file created successfully and content added"


} else {

    Write-Output "Folder already exists."

}



if($flag -eq 0)
{
    if(($name) -and
           ($userLogonName) -and
           ($fullName) -and
           ($department) -and
           ($manager ) -and
           ($title )
           )
    {

    try{
   
        New-ADUser -Name $fullName -SamAccountName $userLogonName1 -GivenName $firstName -Surname $lastName -DisplayName $fullName -Organization $organization -UserPrincipalName ("{0}$userLogonNamedomain" -f $userLogonName) -Office $location -OtherAttributes @{'ExtensionAttribute11'="OFFICE_USER" 
    'ExtensionAttribute1'="DNKHR"
    'ExtensionAttribute13'="VKR-OTH"
    #'manager'=$manager
    'description'=$fullName
    'department'=$department
    'title'=$title 
    'company'=$organization} -AccountPassword (ConvertTo-SecureString -AsPlainText $Password -Force ) -Enabled $true -ChangePasswordAtLogon $true -Path $ouPath -Credential $cred -ErrorAction Stop 
        $worknotes+=Set-IxWorkNotesText -Flag SUCCESS -Note "User [$(Format-IxText $name)](SAMACCOUNTNAME : $userLogonName1) added to AD successfully"

    }
    catch{
        Write-Host $_
        $worknotes+=Set-IxWorkNotesText -Flag ERROR -Note "Unable to create [$(Format-IxText $name)](SAMACCOUNTNAME : $userLogonName1) to the ad user"
        $flag="2"
    }

    if($flag -eq 0)
    {
        Start-Sleep -Seconds 180
        try{
            $Ad_created = Get-ADUser -Identity $userLogonName1 -ErrorAction stop  
            try{
                  Set-ADUser -Identity $userLogonName1 -EmailAddress $email -Credential $cred

                }catch{
                    Write-Host $_
                    $worknotes+= Set-IxWorkNotesText -Flag ERROR -Note "Failed to update the Email address for user : [$(Format-IxText $userLogonName1)]" ##WARNING
                    $flag="2"
                }

            
            
        }
        catch{
            Write-Host $_
            $Ad_created = $null
            $flag="2"
        }
        if($Ad_created -ne $null -and $flag -eq "0")
        {
            
            $worknotes+= Set-IxWorkNotesText -Flag SUCCESS -Note "[POST:CHECK]: New User created [$(Format-IxText $name)](SAMACCOUNTNAME : $userLogonName1) Successfully"
            try{
                $worknotes += set-userproxy -userLogonName1 $userLogonName1
            }catch{
                Write-Host $_        
                $worknotes += Set-IxWorkNotesText -Flag ERROR -Note "Unable to set the proxyaddresses"
                $flag="2"
            }

            try{
            #ask doubt
                $worknotes += sendmailpassword -To_address $ext_Email
            }catch{
                Write-Host $_        
                $worknotes += Set-IxWorkNotesText -Flag ERROR -Note "Unable to send the Password"
                $flag="2"
            }
            
            try{
                $worknotes += sendmaildetails -To_address $ext_Email
                $flag="1"
            }catch{
                Write-Host $_        
                $worknotes += Set-IxWorkNotesText -Flag ERROR -Note "Unable to send the details"
                $flag="2"
            }
        }else{
            $worknotes += Set-IxWorkNotesText -Flag ERROR -Note "User [$(Format-IxText $userLogonName1)] is not found in AD.Possibly Sync issue."
        }
    }
    }
    else{
        $worknotes+= Set-IxWorkNotesText -Flag ERROR -Note "Some mandate Inputs are missing"
        $flag = "2"
    }
}

if($flag -eq "1" -and $global:tempflag -eq "0"){
    try{
        Update-IxTicket -sys_id $sys_id -table $table -work_notes $worknotes -case "Success"
        #Update-IxMongoData -id $id -status Completed
    }
    catch{
        "Error occured while updating worknotes in service now"
    }
}

else{
    try{
        Update-IxTicket -sys_id $sys_id -table $table -work_notes $worknotes -case "failed" -assignment_group $assignment_group
        #Update-IxMongoData -id $id -status Failed
        write-Host "failed"
    }
    catch{
        $_
        "Error occured while updating worknotes in service now"
    }
}

}


elseif($short_description -like "*Assign relevant O365 licenses for the requested external user*"){



if ($textFilePath){
    $folderPath = "C:\temp\$Ritm"
    $textFilePath="$folderPath\$Ritm.txt"
    #$logonNameofUser = (Get-Content -Path $textFilePath).Trim()
    $fileContent = Get-Content -Path $textFilePath
  
    $logonNameofUser = ($fileContent[0]).Trim()

    $userLogonName = ($fileContent[1]).Trim()

    $userLogonDomainName = ($fileContent[2]).Trim()
}
else{
Write-Host "The folder not found"
}
 

    try{

    $groups=@(
    "O365-Support-License-M365-E5",
    "NetAccess-Internet-VEL-Users",
    "NetAccess-Internet-Extended-VEL-Users",
    "NetAccess-Mobile-devices-Users-ModernMail",
    "NetAccess-NPS-VPN-Users")

    $UserOU=(Get-ADUser $logonNameofUser -Properties Member).DistinguishedName

    foreach ($group in $groups) {
        #Get-ADGroup $group -Properties Member).Member -contains $UserOU

            #$isMember = Get-ADGroupMember -Identity $group | Where-Object { $_.SamAccountName -eq $userLogonName1 }
            $isMember = (Get-ADGroup $group -Properties Member).Member -contains $UserOU



        if ($isMember) {
            Write-Output "$logonNameofUser is already a member of $group"
            $worknotes += Set-IxWorkNotesText -Flag Info -Note "$logonNameofUser is already a member of $group"
        }
         else {
            Add-ADGroupMember -Identity $group -Members $logonNameofUser
            $userInGroup_PostCheck = (Get-ADGroup $group -Properties Member).Member -contains $UserOU }

            if ($userInGroup_PostCheck ) {
                Write-Output "$logonNameofUser is successfully added to $group"
                $worknotes += Set-IxWorkNotesText -Flag Success -Note "$logonNameofUser is successfully added to $group"
            }
            else{
                Write-Output "Unable to add $logonNameofUser to $group"
                $worknotes += Set-IxWorkNotesText -Flag Error -Note "Unable to add $logonNameofUser to $group"
                $flag = 2
            }

        }      
    

    }catch{
        write-host $_
        $worknotes += Set-IxWorkNotesText -Flag ERROR -Note "Error While Providing license to  $logonNameofUser"
        $flag = 2

    }




    if($flag -eq "1" -and $global:tempflag -eq "0"){
    try{
        Update-IxTicket -sys_id $sys_id -table $table -work_notes $worknotes -case "Success"
        #Update-IxMongoData -id $id -status Completed
    }
    catch{
        "Error occured while updating worknotes in service now"
    }
}

    else{
    try{
        #Update-IxTicket -sys_id $sys_id -table $table -work_notes $worknotes -case "failed" -assignment_group $assignment_group
        #Update-IxMongoData -id $id -status Failed
        write-Host "failed"
    }
    catch{
        $_
        "Error occured while updating worknotes in service now"
    }
}



    #Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
    #Import-Module ExchangeOnlineManagement




Write-Host "Waiting for 40 minutes before executing the next script..." -ForegroundColor Yellow
Start-Sleep -Seconds 2400

Write-Host "40 minutes completed. Running the second script now..." -ForegroundColor Green

Connect-ExchangeOnline -Credential $cred

$users ="$userLogonName$userLogonNamedomain"

$fileDeletion=1

try{

foreach ($user in $users)
{
$Userprincipalname = $user
Enable-Mailbox $Userprincipalname -Archive
$Calendar = "$userprincipalname" + ":\calendar"
Set-Mailbox $Userprincipalname -RetentionPolicy "VELUX Default User MRM Policy"
Set-Mailboxfolderpermission $calendar -User Default -Accessrights Limiteddetails

}

}catch{
$fileDeletion=2
}


if($fileDeletion -eq "1"){
Remove-Item -Path $folderPath -Recurse -Force 
Write-Output "Folder and all contents deleted successfully."
}
else{
Write-Host "Unable to delete the file"
}
 
}