function getMail($nihuz, $id)
{
    $nihuz.objects | ?{$_.id -eq $id} | select -ExpandProperty mail
}

$path = "C:\Hackathon\Inbox"
$sent = "C:\Hackathon\Sent"
$errors = "C:\Hackathon\Errors"
$logs = "C:\Hackathon\hackathonLog.txt"


(Get-Date).ToString() + " Started import flow" >> $logs

try{
    $nihuz = Get-Content C:\Hackathon\nihuz.json | ConvertFrom-Json -ErrorAction SilentlyContinue
    $items = Get-ChildItem -Path $path

    $items | %{
    
        $file = $_
        try{
            $xml =  [XML] ($file |Get-Content)

            $subject = $xml.ChildNodes.subject
            $body = $xml.ChildNodes.body
            $from = getMail $nihuz $xml.ChildNodes.from
    
            $recipients = $xml.ChildNodes.to.Split(";") | %{getMail $nihuz $_.trim() }

            try{
                Send-MailMessage -SmtpServer xch.dsam.test -Subject $subject -Body $body -To $recipients -From $from
                Move-Item $file.fullname -Destination $sent
                (Get-Date).ToString() + " Sent $($_.name)"  >> $logs

            }
            catch{
                Move-Item $file.FullName -Destination $errors
                (Get-Date).ToString() + " Error sending $($file.name)"
            }
        }
        catch{
            (Get-Date).ToString() + " Failed to import mail xml $($file.name)"  >> $logs
            Move-Item $file.FullName -Destination $errors
        }
        
            
    }

}
catch{
    (Get-Date).ToString() + " Import from Nihuz failed"  >> $logs
}