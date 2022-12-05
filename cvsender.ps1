############################################################################
$mailovka = "michal@kopl.pro"; #must be configured in Your Outlook
$LogFile = "C:\_DEV\cv-sender\log.txt"
$output_path_tmp = "C:\_DEV\cv-sender\tmp\adresy.txt"
$output_path = "C:\_DEV\cv-sender\adresy.txt"
$SenderName = "Michal Kopl"
$SenderMail = "michal@kopl.pro"
$SenderMobile = "+420 777 062 425"
$SenderMobileLink = $SenderMobile.Trim(" ")
$CVLink = "https://www.kopl.pro/images/Michal-Kopl_CV.pdf"
$CertificationsLink = "https://kopl.pro/images/Michal-Kopl-certifikace.rar"
$WebLink = "https://www.kopl.pro"
$LinkedInLink = "https://www.linkedin.com/in/michalkopl/"
$GitHubLink = "https://github.com/michal-kopl"
$CVFile = "C:\_DEV\cv-sender\Michal-Kopl_CV.pdf"
$DigitalReportFile = "C:\_DEV\cv-sender\Michal-Kopl_Digital-Skills-Report.pdf"
$mysql_server = "localhost"
$mysql_user = "root"
$mysql_password = ""
$dbName = "kontakty" 
$frequency=30
#$unit = "MINUTE" #Debugging
$unit = "DAY" #Production
[void][system.reflection.Assembly]::LoadFrom("C:\_DEV\cv-sender\dll\MySql.Data.dll") 
############################################################################
function Invoke-SetProperty {
    param(
        [__ComObject] $Object,
        [String] $Property,
        $Value
    )
    [Void] $Object.GetType().InvokeMember($Property,"SetProperty",$NULL,$Object,$Value)
}
$Connection = New-Object -TypeName MySql.Data.MySqlClient.MySqlConnection
$Connection.ConnectionString = "SERVER=$mysql_server;DATABASE=$dbName;UID=$mysql_user;PWD=$mysql_password"
$Connection.Open()
$sql = New-Object MySql.Data.MySqlClient.MySqlCommand
$sql.Connection = $Connection
$sql.CommandText = "SELECT ``email``, ``last_message``, ``use`` FROM ``emaily_personalky`` WHERE ``use``=1 AND ``last_message``<ADDDATE(``last_message``,INTERVAL $frequency $unit) AND TIMESTAMP(NOW())>ADDDATE(``last_message``,INTERVAL $frequency $unit);" 
Write-Host $sql.CommandText
$dataAdapter = New-Object MySql.Data.MySqlClient.MySqlDataAdapter($sql)
$dataSet = New-Object System.Data.DataSet
$dataAdapter.Fill($dataSet) | Out-Null
$dataSet.Tables[0] | Export-Csv -path $output_path_tmp -NoTypeInformation
Get-Content $output_path_tmp | ForEach-Object {$_ -replace '"', ''} | Select-Object -Skip 1 | out-file $output_path -Fo -En ascii
$Subject = "Seeking for a new job in IT - $SenderName";
$htmltext = "<!DOCTYPE html><html style='font-size: 16px;'> <head> <meta name='viewport' content='width=device-width, initial-scale=1.0'> <meta charset='utf-8'> <meta name='keywords' content=''> <meta name='description' content=''> <meta name='page_type' content='np-template-header-footer-from-plugin'> <title>Home</title> <link id='u-theme-google-font' rel='stylesheet' href='https://fonts.googleapis.com/css?family=Roboto:100,100i,300,300i,400,400i,500,500i,700,700i,900,900i|Open+Sans:300,300i,400,400i,500,500i,600,600i,700,700i,800,800i'> <script type='application/ld+json'>{'@context': 'http://schema.org','@type': 'Organization','name': '$SenderName'}</script> <meta name='theme-color' content='#478ac9'> <meta property='og:title' content='About $SenderName'> <meta property='og:description' content='Searching for new post related to my previous experiences.'> <meta property='og:type' content='website'> <STYLE> tr{text-align: center}td{text-align: center;padding:20px}th{text-align: center;padding:20px}table{text-align: center; width:600;margin:auto;margin-top:90px}a{text-decoration: underline;text-decoration-color: cornflowerblue;}.background{background-image: linear-gradient(270deg, #f5f7fa, #eeeeee);}</STYLE> </head> <body class='u-body u-gradient u-xl-mode' style='background-image: linear-gradient(270deg, #f5f7fa, #b3b3b3);'> <CENTER><TABLE border='0' style='width:600px;'><TR><TD colspan='2' align='center'style='background-image: linear-gradient(270deg, #f5f7fa, #b3b3b3);'><H1>Seeking for a new position in IT scope</H1><br /></TD></TR> <TR><TD colspan='2' class='background'><h2>About Me</h2><p style='text-align:justify'>I do have more than 22 years of working experience in scope of IT, management of OS (MS Windows, LINUX), networking and webdesign. I also provided support to end users and management, creating documentation and Search Engine Optimization with intent to make the website presentation visible and providing customers' flow.</p><TABLE border='0' style='width:600px'>
<TR>
    <TD><a href='$CVLink'><svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 52 52' width='35' height='35'><g fill='none' stroke='#000' stroke-linecap='round' stroke-linejoin='round' stroke-miterlimit='10' stroke-width='2'><path d='M30.073 21.444h7.886M14.041 28.783H37.96M14.041 36.123H37.96M14.041 43.462H37.96M24.073 23.244H12.547V20.88c0-2.7 2.188-4.888 4.888-4.888h1.75c2.7 0 4.888 2.188 4.888 4.888v2.364z'/><circle cx='18.31' cy='10.001' r='3'/><path d='M26.323 14.575H37.96M33.65 2H9.201a2.15 2.15 0 00-2.15 2.15v43.7A2.15 2.15 0 009.202 50h33.596a2.15 2.15 0 002.15-2.15V13.3a2.15 2.15 0 00-.63-1.52l-9.149-9.15A2.15 2.15 0 0033.65 2z'/></g></svg><br />My CV</a></TD>
    <TD><a href='$CertificationsLink'><svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 52 52' width='35' height='35'><g fill='none' stroke='#000' stroke-linecap='round' stroke-linejoin='round' stroke-miterlimit='10' stroke-width='2'><path d='M36.813 22.937v-9.97M31.333 40.907H8.983c-1.06 0-1.91-.85-1.91-1.91V3.217c0-.75.61-1.36 1.36-1.36h16.74'/><path d='M25.176 1.853v9.757c0 .751.61 1.36 1.36 1.36h10.279L25.176 1.853zM10.689 10.143h9.945M10.689 18.362h20.306M10.689 24.593h19.766M26.733 30.827h-16.04M44.927 32.028c0 3.54-2.01 6.611-4.952 8.123a8.477 8.477 0 01-1.024.462 9.45 9.45 0 01-2.674.553c-.156.009-.322.009-.48.009h-.018c-.36 0-.719-.018-1.07-.065a8.834 8.834 0 01-2.6-.71 5.18 5.18 0 01-.507-.249 9.116 9.116 0 01-4.96-8.123c0-5.053 4.093-9.148 9.137-9.148a9.146 9.146 0 019.148 9.148z'/><path d='M31.603 40.147l-.27.76-3.22 9.24h5.78l1.88-8.97M39.973 40.147l3.48 10h-5.78l-1.86-8.97'/></g></svg><br />My Certifications</a></TD>
    <TD><a href='$WebLink'><svg xmlns='http://www.w3.org/2000/svg' width='35' height='35'><path d='M57.334 46.09a28.91 28.91 0 0 0 0-28.18.971.971 0 0 0-.406-.706 29 29 0 1 0 0 29.592.971.971 0 0 0 .406-.706ZM15.38 47h4.579A27.377 27.377 0 0 0 25.1 57.342 22.448 22.448 0 0 1 15.38 47ZM5.025 33h4.994a36.32 36.32 0 0 0 2.324 12h-4a26.807 26.807 0 0 1-3.318-12Zm3.32-14h4a36.32 36.32 0 0 0-2.324 12H5.025a26.807 26.807 0 0 1 3.32-12Zm40.275-2h-4.579A27.377 27.377 0 0 0 38.9 6.658 22.448 22.448 0 0 1 48.62 17Zm10.355 14h-4.994a36.32 36.32 0 0 0-2.324-12h4a26.807 26.807 0 0 1 3.318 12Zm-6.994 0h-5.993a56.232 56.232 0 0 0-1.427-12h4.961a34.255 34.255 0 0 1 2.459 12ZM31 19v12H20.012a53.521 53.521 0 0 1 1.526-12Zm-8.91-2c2-6.586 5.252-11.177 8.91-11.894V17ZM31 33v12h-9.462a53.521 53.521 0 0 1-1.526-12Zm0 14v11.894c-3.658-.717-6.909-5.308-8.91-11.894Zm2-2V33h10.988a53.521 53.521 0 0 1-1.526 12Zm8.91 2c-2 6.586-5.253 11.177-8.91 11.894V47ZM33 31V19h9.462a53.521 53.521 0 0 1 1.526 12Zm0-14V5.106c3.657.717 6.908 5.308 8.91 11.894ZM25.1 6.658A27.377 27.377 0 0 0 19.959 17H15.38A22.448 22.448 0 0 1 25.1 6.658ZM19.438 19a56.234 56.234 0 0 0-1.426 12h-5.993a34.255 34.255 0 0 1 2.458-12Zm-7.419 14h5.993a56.234 56.234 0 0 0 1.426 12h-4.961a34.255 34.255 0 0 1-2.458-12ZM38.9 57.342A27.377 27.377 0 0 0 44.041 47h4.579a22.448 22.448 0 0 1-9.72 10.342ZM44.561 45a56.232 56.232 0 0 0 1.427-12h5.993a34.255 34.255 0 0 1-2.459 12Zm9.42-12h4.994a26.807 26.807 0 0 1-3.32 12h-4a36.32 36.32 0 0 0 2.326-12Zm.456-16h-3.615a27.44 27.44 0 0 0-6.5-9.012A27.187 27.187 0 0 1 54.437 17Zm-34.76-9.012A27.452 27.452 0 0 0 13.178 17H9.563a27.184 27.184 0 0 1 10.114-9.012ZM9.563 47h3.615a27.452 27.452 0 0 0 6.5 9.012A27.184 27.184 0 0 1 9.563 47Zm34.759 9.012a27.44 27.44 0 0 0 6.5-9.012h3.615a27.187 27.187 0 0 1-10.115 9.012Z'/></svg><br />My Website</a></TD>
    <TD><a href='$LinkedInLink'><span ng-if='showField('linkedinURL')' class='ng-scope'><a href='$LinkedInLink' target='_blank' rel='noopener'><svg fill='#000000' xmlns='http://www.w3.org/2000/svg'  viewBox='0 0 50 50' width='35px' height='35px'><path d='M25,2C12.3,2,2,12.3,2,25s10.3,23,23,23s23-10.3,23-23S37.7,2,25,2z M19,35c0,0.5-0.5,1-1,1h-4c-0.5,0-1-0.5-1-1V20 c0-0.5,0.5-1,1-1h4c0.5,0,1,0.5,1,1V35z M16,18c-1.6,0-3-1.4-3-3s1.4-3,3-3s3,1.4,3,3S17.6,18,16,18z M38,35c0,0.5-0.5,1-1,1h-4 c-0.5,0-1-0.5-1-1v-7.5c0-1.4-1.1-2.5-2.5-2.5S27,26.1,27,27.5V35c0,0.5-0.5,1-1,1h-4c-0.5,0-1-0.5-1-1V20c0-0.5,0.5-1,1-1h4 c0.5,0,1,0.4,1,1c1.1-0.6,2.2-1,3.5-1c4.1,0,7.5,3.4,7.5,7.5V35z'/></svg><br />My Linked In profile</a></a></TD>
    <TD><a href='$GitHubLink' data-hotkey='g d' aria-label='Homepage ' data-turbo='false' data-analytics-event='{&quot;category&quot;:&quot;Header&quot;,&quot;action&quot;:&quot;go to dashboard&quot;,&quot;label&quot;:&quot;icon:logo&quot;}'><svg height='35' aria-hidden='true' viewBox='0 0 16 16' version='1.1' width='35' height='35' data-view-component='true' class='octicon octicon-mark-github v-align-middle'><path fill-rule='evenodd' d='M8 0C3.58 0 0 3.58 0 8c0 3.54 2.29 6.53 5.47 7.59.4.07.55-.17.55-.38 0-.19-.01-.82-.01-1.49-2.01.37-2.53-.49-2.69-.94-.09-.23-.48-.94-.82-1.13-.28-.15-.68-.52-.01-.53.63-.01 1.08.58 1.23.82.72 1.21 1.87.87 2.33.66.07-.52.28-.87.51-1.07-1.78-.2-3.64-.89-3.64-3.95 0-.87.31-1.59.82-2.15-.08-.2-.36-1.02.08-2.12 0 0 .67-.21 2.2.82.64-.18 1.32-.27 2-.27.68 0 1.36.09 2 .27 1.53-1.04 2.2-.82 2.2-.82.44 1.1.16 1.92.08 2.12.51.56.82 1.27.82 2.15 0 3.07-1.87 3.75-3.65 3.95.29.25.54.73.54 1.48 0 1.07-.01 1.93-.01 2.2 0 .21.15.46.55.38A8.013 8.013 0 0016 8c0-4.42-3.58-8-8-8z'></path></svg><br />My GitHub</a></TD>
</TR>
<TR><TD colspan=4>Mobile phone: <a href='tel:$SenderMobileLink'>$SenderMobile</a><br />email: <a href='mailto:$SenderMail'>$SenderMail</a></TD></TR></TABLE><BR />Looking forward to your reply with eventual employment vacancy.<BR /><BR />Best regards,<BR />$SenderName</TD></TR></TABLE></CENTER></body></html>";
Get-Content -Path $output_path | ForEach-Object -Process {
    #$sleeping = 1;
    $sleeping = Get-Random -Minimum 15 -Maximum 55;
    $Adresa = $_;
    $CharArray = $Adresa.Split(",");$mail=$CharArray['0'].Trim();
    $outlook = new-object -comobject outlook.application;
    $email = $outlook.CreateItem(0);
    $account = $email.Session.Accounts.Item($mailovka);
    Invoke-SetProperty -Object $email -Property "SendUsingAccount" -Value $account;
    $email.To = $mail;
    $email.Subject = $Subject;
    $priloha=$CharArray['2'].Trim();
    if($priloha -eq "1"){
        $email.Attachments.Add($CVFile);
        $email.Attachments.Add($DigitalReportFile);
    };
    $email.HTMLBody = $htmltext;
    Start-Sleep $sleeping;
    Write-Output "email: $mail|$sleeping";    
    $datetimestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $email.Send();
    $sql.CommandText = "UPDATE ``emaily_personalky`` SET ``last_message``='$datetimestamp' WHERE ``email``='$mail';"
    $dataAdapter = New-Object MySql.Data.MySqlClient.MySqlDataAdapter($sql)
    $dataSet = New-Object System.Data.DataSet
    $dataAdapter.Fill($dataSet) | Out-Null
    $mail | out-file -filepath $LogFile      #last e-mail message recipient's address stored within
    #$email.save(); #//for creating drafts in Outlook only
}
$Connection.Close()