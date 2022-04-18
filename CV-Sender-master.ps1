$mailovka = "first.name@domain.net"; #have to be configured in Your Outlook
function Invoke-SetProperty {
    param(
        [__ComObject] $Object,
        [String] $Property,
        $Value
    )
    [Void] $Object.GetType().InvokeMember($Property,"SetProperty",$NULL,$Object,$Value)
}
$Subject = "Seeking for a new job post - Firstname Surname";
$htmltext = "<!DOCTYPE html><html style='font-size: 16px;'> <head> <meta name='viewport' content='width=device-width, initial-scale=1.0'> <meta charset='utf-8'> <meta name='keywords' content=''> <meta name='description' content=''> <meta name='page_type' content='np-template-header-footer-from-plugin'> <title>Home</title> <link id='u-theme-google-font' rel='stylesheet' href='https://fonts.googleapis.com/css?family=Roboto:100,100i,300,300i,400,400i,500,500i,700,700i,900,900i|Open+Sans:300,300i,400,400i,500,500i,600,600i,700,700i,800,800i'> <script type='application/ld+json'>{'@context': 'http://schema.org','@type': 'Organization','name': 'Firstname Surname'}</script> <meta name='theme-color' content='#478ac9'> <meta property='og:title' content='About Firstname Surname'> <meta property='og:description' content='Searching for new post related to my previous experiences.'> <meta property='og:type' content='website'> <STYLE> tr{text-align: center}td{text-align: center;padding:20px}th{text-align: center;padding:20px}table{text-align: center; margin:auto;margin-top:90px}a{text-decoration: underline;text-decoration-color: cornflowerblue;}.background{background-image: linear-gradient(270deg, #f5f7fa, #eeeeee);}</STYLE> </head> <body class='u-body u-gradient u-xl-mode' style='background-image: linear-gradient(270deg, #f5f7fa, #b3b3b3);'> <CENTER><TABLE border='0' Width='600px'> <TR><TD colspan='2' align='center'style='background-image: linear-gradient(270deg, #f5f7fa, #b3b3b3);'><H1>Seeking for a new position in IT scope</H1><br /></TD></TR> <TR><TD colspan='2' class='background'><h2>About Me</h2><p style='text-align:justify'>I do have more than 22 years of working experience in scope of IT, networking and webdesign. I also provided a support to end users and management, creating documentation and also Search Engine Optimization with intent to make the website presentation visible and providing customers' flow.</p><TABLE border='0' width='600px'><TR><TD><a href='https://www.seznam.cz'>My CV</a></TD><TD><a href='https://www.seznam.cz/download/Certifikcates.rar'>Certificates</a></TD><TD><a href='https://www.seznam.cz/#contact-details'>Personal website</a></TD><TD><a href='https://www.linkedin.com/in/first.name/'>Linked-In</a></TD></TR><TR><TD colspan=4>Mobile phone: <a href='tel:+420111222333'>+420 111 222 333</a><br />email: <a href='mailto:first.name@domain.com'>first.name@domain.com</a></TD></TR></TABLE><BR />Best regards,<BR />Firstname Surname</TD></TR></TABLE><BR/></CENTER></body></html>";
Get-Content -Path "C:\_DEV\CVposting\adresy.txt"|ForEach-Object -Process {
    $sleeping = Get-Random -Minimum 45 -Maximum 75;
    $Adresa = $_;
    $CharArray = $Adresa.Split(",");$mail=$CharArray['0'].Trim();
    #$osloveni=$CharArray['1'].Trim();
    $outlook = new-object -comobject outlook.application;
    $email = $outlook.CreateItem(0);
    $account = $email.Session.Accounts.Item($mailovka);
    Invoke-SetProperty -Object $email -Property "SendUsingAccount" -Value $account;
    $email.To = $mail;
    $email.Subject = $Subject;
    $priloha=$CharArray['2'].Trim();
    if($priloha -eq "0"){
        #-- no attachment added"
    }else {
        $email.Attachments.Add("C:\_DEV\CVposting\CV-Firstname-Surname.pdf");
    };
    $email.HTMLBody = $htmltext;
    Start-Sleep $sleeping;
    Write-Output "email: $mail|$sleeping";    
    $email.Send(); #$email.save(); //for creating drafts in Outlook only
}