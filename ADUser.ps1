Param($id)
#$id="66fb86cbaba653c9d22a0e0a"

$flag=0
$global:tempflag=0

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
$title                 = $data.validated_inputs.'parent.variables.title'
$organization          = $data.validated_inputs.'parent.variables.which_organization_are_the_user_to_work_for'
$location              = $data.validated_inputs.'parent.variables.location'
$manager               = $data.validated_inputs.'parent.variables.manager'
$department            = $data.validated_inputs.'parent.variables.departement'
$shortname             = $data.validated_inputs.'parent.variables.requested_shortname'
$ext_Email           = $data.validated_inputs.'parent.variables.email'


$sys_id                = $data.validated_inputs.sys_id
$table                 = "sc_task"
$assignment_group      = $data.fulfillment_group_sysid

$cred = Get-IxServiceAccount
$worknote=Set-IxWorkNotesText "Automation work in progess"
try{Update-IxTicket -sys_id $sys_id -table $table -work_notes $worknote -case "wip"}catch{$_}
try{Update-IxMongoData -id $id -status WorkInProgress}catch{$_}

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
$fullName = "$firstName $lastName"

$OUpath = ""

if ($Organization -eq "VKR Holding A/S") {
    $OUpath = 'OU=Standard Users,OU=Users,OU=DNKHR,OU=Locations,DC=velux,DC=org'
} else {
    $OUpath = 'OU=Standard Users,OU=Users,OU=DNKSR,OU=Locations,DC=velux,DC=org'
}




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
        $userLogonNamedomain = "@veluxfoundations.dk"
    }
    "ADDEK" {
        $domainName = ".vkr"
        $userLogonNamedomain = "@velux.com"
    }
    "Fondene" {
        $domainName = ".fond"
        $userLogonNamedomain = "@velux.com"
    }
    default {
        Write-Host "Unknown organization"
    }
}

function Generate-InternalUserID{
    Param($name)
    $name = $name -replace '[^a-zA-Z]', ''
    foreach($index in $(0..$($name.Length-1))){
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


$userLogonName = ($userLogonName1 -split '\.')[0] 



function set-userproxy

{

    param([string] $User_id)
 
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
 
    try{


        $currentUser = Get-ADUser -Identity $User_id -Properties proxyAddresses

        $existingProxyAddresses = $currentUser.proxyAddresses | Where-Object { 

            $_ -notmatch $userLogonNamedomain

        }
 

        $finalProxyAddresses = $existingProxyAddresses + $proxyaddresses + $newProxyAddresses
 

        Set-ADUser -Identity $User_id -Replace @{

            mail = $mail;

            mailNickname = $mailnickname;

            targetAddress = $targetaddress;

            proxyAddresses = $finalProxyAddresses

        }
 
        $setproxy = Set-IxWorkNotesText -Flag SUCCESS -Note "Successfully updated proxy_addresses for User : [$(Format-IxText $User_id)]"

    }catch{

        Write-Host $_

        $setproxy = Set-IxWorkNotesText -Flag ERROR -Note "Failed to update the proxy address for User : [$(Format-IxText $User_id)]"

        $global:tempflag=1

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
                $worknotes += set-userproxy -User_id $userLogonName1
            }catch{
                Write-Host $_        
                $worknotes += Set-IxWorkNotesText -Flag ERROR -Note "Unable to set the proxyaddresses"
                $flag="2"
            }

            try{
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
        Update-IxMongoData -id $id -status Completed
    }
    catch{
        "Error occured while updating worknotes in service now"
    }
}else{
    try{
        Update-IxTicket -sys_id $sys_id -table $table -work_notes $worknotes -case "failed" -assignment_group $assignment_group
        Update-IxMongoData -id $id -status Failed
        write-Host "failed"
    }
    catch{
        $_
        "Error occured while updating worknotes in service now"
    }
}