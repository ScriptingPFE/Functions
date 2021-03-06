Function Send-OutlookEmail {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false, 
            Position = 0)]    
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]    
        [String[]]$To,

        [Parameter(Mandatory = $false, 
            Position = 1)]    
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]    
        [String]$Bcc,

        [Parameter(Mandatory = $false, 
            Position = 2)]    
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]    
        [String]$Subject,

        [Parameter(Mandatory = $false, 
            Position = 3)]    
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]    
        [String]$Body = '',

        [Parameter(Mandatory = $false, 
            Position = 4)]    
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]    
        [Switch]$BodyIsHTML,

        [Parameter(Mandatory = $false, 
            Position = 5)]    
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]    
        [ValidateSet("Low", "Normal", "High")]
        $MessagePriority,
        
        [Parameter(Mandatory = $false, 
            Position = 6)]    
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]    
        [switch]$RequestReadReceipt

    )
    if ($to.count -eq 0 -and $bcc.count -eq 0) {
        Write-Error -Message "A recipient is required to send a message." -RecommendedAction "Specify a recipient for either the to or Bcc parameter"
    }
    Else {
        $outlook = New-Object -comObject Outlook.Application 
        $Mail = $Outlook.CreateItem(0)
        $Mail.Subject = [String]$Subject
        $mail.BCC = $Bcc

        Foreach ($Recipient in $To) {
            $Mail.Recipients.Add($Recipient) | Out-Null
        }

        if ($BodyIsHTML) {
            $mail.HTMLBody = [string]($Body) 
        }
        Else {
            $mail.Body = [string]($Body)
        }

        if($RequestReadReceipt){
            $mail.ReadReceiptRequested = $true
        }
        
        switch ($MessagePriority) {
            Low { $mail.Importance = 0 ; Break }
            Normal { $mail.Importance = 1 ; Break }
            High { $mail.Importance = 2 ; Break }
        }


        $Mail.Send()
        #$Outlook.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null
    }
}
