Function Get-O365UrlAvailability {
[cmdletbinding()]
    param(
        [Parameter(Mandatory = $true,
        ValueFromPipelineByPropertyName = $true,
        Position = 0)] 
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("USDOD","GCCHIGH","COMMERCIAL","Germany","CHINA")]
        $O365TentantNetwork,
        [Parameter(Mandatory = $true,
        ValueFromPipelineByPropertyName = $true,
        Position = 0)] 
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [ValidatePattern('(\w+@[a-zA-Z_]+?\.[a-zA-Z]{2,6})')]
        $EmailAddress

    )
    <#
    .SYNOPSIS
        NAME: Get-O365UrlAvailability
        AUTHOR: Eric Powers ericpow@microsoft.com
        VERSION: 2.1

    .FUNCTIONALITY
         

    .DESCRIPTION
        Attempts to make connections with each of the IPs for the O365 endpoints based on the selection provided in the O365TentantNetwork parameter. 
        Requires the file from O365 Endpoints Json which is downloaded for the endpoint you select automatically
        
        To download this file programatically you can use the lines below
        $webclient = [System.Net.WebClient]::new()
        $webclient.DownloadFile('https://endpoints.office.com/endpoints/USGOVDoD?clientrequestid=b10c5ed1-bad1-445f-b386-b919946339a7',"$env:USERPROFILE\Desktop\O365USDODNetworkRequirements.json")


    .EXAMPLE
        Get-O365UrlAvailability.ps1 -O365TentantNetwork Commercial -EmailAddress ericpow@microsoft.com

    .NOTES
        20210107: v1.0 - Initial Release
        20210407: v2.0 - Updated to validate url endpoints only
        0210407: v2.1 - Updated to validate url endpoints only use emailaddress as well as O365TentantNetwork to select the proper endpoint for public consumption
        
#>
Function Get-NetworkConnectionAvailablitiy {
    Param($Computername,
    $Port
    )
        $stopwatch = new-object system.diagnostics.stopwatch
        $stopwatch.Reset()
        $stopwatch.Start()
    
        $Cntr =  0
        $TCPClient= New-object System.Net.Sockets.TcpClient
            if( ($TCPClient | Get-Member -MemberType Method | Select-Object -ExpandProperty Name) -contains 'ConnectAsync'){
                $Result = $TCPClient.ConnectAsync($Computername,$Port) 
                do{Start-Sleep -seconds 1;$Cntr++}Until(($Result.IsCompleted -eq $true -or $Cntr -ge 5) )
            }
            Else{
                $Result = $TCPClient.Connect($Computername,$Port) 
                do{Start-Sleep -seconds 1;$Cntr++}Until(($Result.IsCompleted -eq $true -or $Cntr -ge 5))
            }
        $stopwatch.stop()
    
        $successOnConnect=if($Result.IsFaulted -or $cntr -ge 5){$False}
        Else{$true}
        New-Object Psobject -Property @{
        SourceHost= $($Env:computername +" " + [string]$LocalIP.IPAddress) ;
        RemoteIPAddress=$([string]$Computername);
        Port = $port; 
        IsAvailable = $successOnConnect;
        testDuration= $([math]::Round($stopwatch.Elapsed.TotalSeconds,2))
        
        } | Select-Object SourceHost,RemoteIPAddress,Port,testDuration,IsAvailable
    
        if($TCPClient.Connected){
            $TCPClient.Close()
            $TCPClient.Dispose()
        }
        Else{
            $TCPClient.Dispose()
        }
    }
    Function Get-RootPrimarySMTPDomain {
        param(        
            $EmailAddress
        )
        (($EmailAddress -split "@")[1] -split "\.")[0]
    }
    
    if(!(Test-path $env:USERPROFILE\Desktop\)){
        new-item $env:USERPROFILE\Desktop\O365Networks -ItemType Directory  | Out-Null
    }
    
    $EmailDomain 
    $webclient = [System.Net.WebClient]::new()
    switch ($O365TentantNetwork) {
        "USDOD" {  
            $webclient.DownloadFile('https://endpoints.office.com/endpoints/USGOVDoD?clientrequestid=b10c5ed1-bad1-445f-b386-b919946339a7',"$env:USERPROFILE\Desktop\O365NetworkRequirements.json")
            break
        } 
        "GCCHIGH" {  
            $webclient.DownloadFile('https://endpoints.office.com/endpoints/USGOVGCCHigh?clientrequestid=b10c5ed1-bad1-445f-b386-b919946339a7', "$env:USERPROFILE\Desktop\O365NetworkRequirements.json")
            break
        } 
        "COMMERCIAL" {  
            $webclient.DownloadFile('https://endpoints.office.com/endpoints/worldwide?clientrequestid=b10c5ed1-bad1-445f-b386-b919946339a7', "$env:USERPROFILE\Desktop\O365NetworkRequirements.json")
            break
        } 
        "Germany" {  
            $webclient.DownloadFile('https://endpoints.office.com/endpoints/Germany?clientrequestid=b10c5ed1-bad1-445f-b386-b919946339a7', "$env:USERPROFILE\Desktop\O365NetworkRequirements.json")
            break
        } 
        "China" {  
            $webclient.DownloadFile('https://endpoints.office.com/endpoints/China?clientrequestid=b10c5ed1-bad1-445f-b386-b919946339a7', "$env:USERPROFILE\Desktop\O365NetworkRequirements.json")
            break
        } 
    }
    
    Write-host -ForegroundColor DarkCyan "Processing O365 Network Requirements File: " -NoNewline; $json
    $json = "$env:USERPROFILE\Desktop\O365NetworkRequirements.json"
    $JsonFile = Get-Content $Json | Convertfrom-Json
    $O365Network =  $O365TentantNetwork
    $SmtpRoot = Get-RootPrimarySMTPDomain -EmailAddress $EmailAddress
    
    
    $findings = Foreach ($JsonEntry in $JsonFile ){
        if(($JsonEntry | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name) -contains 'urls'){
            foreach ($url in $JsonEntry.urls){
                    if([regex]::Match($url ,"\A\*").success){
                        $testedurl = $url.trimstart("*.")
                    }
                    Elseif([regex]::Match($url ,".\*.").success){
    
                        $testedurl = $url -replace "\*",$Smtproot
                     
                    }
                    else{
                        $testedurl = $url 
                    }                    
                    
                    $isAvailable = $True                    
                    foreach ($port in ($JsonEntry.tcpPorts -split ',')){
                        $reportfinding = [pscustomobject]@{
                            id = $JsonEntry.id
                            O365Network = $O365Network
                            category = $JsonEntry.category
                            expressRoute = $JsonEntry.expressRoute
                            serviceArea = [string]($JsonEntry| Select-Object -expandproperty serviceArea) +';'
                            serviceAreaDisplayName = [string]($JsonEntry| Select-Object -expandproperty serviceAreaDisplayName) +';' 
                            required = $JsonEntry.required
                            tcpPorts = $JsonEntry.tcpPorts
                            testedPort= $port
                            requiredURL = $url 
                            testedURL = $testedurl
                            success = $false 
                            testDuration = $null
                            }
                                
                        $connectionResult = Get-NetworkConnectionAvailablitiy -Computername $testedurl  -Port $port -WarningAction SilentlyContinue 
                        $reportfinding.testDuration = $connectionResult.testDuration
    
                            
                        if($reportfinding.testDuration -ge 5){
                            $duration = 'timed out'
                        }
                        Else{
                            $duration = $reportfinding.testDuration
                        }   
    
                        if($connectionResult.IsAvailable -eq $false){
                            $reportfinding.success = $false
                            
                            write-host -ForegroundColor Red "Url: " -NoNewline; Write-Host $TestedUrl -NoNewline;write-host -ForegroundColor Red " TCPPort: " -nonewline; write-host $port -nonewline; write-host -ForegroundColor Red " is NOT available. Test duration (Seconds): " -NoNewline;Write-Host $duration
                        }
                        else{
                            $reportfinding.success = $true
                            write-host -ForegroundColor green "Url: " -NoNewline; Write-Host $TestedUrl -NoNewline;write-host -ForegroundColor green " TCPPort: " -nonewline; write-host $port -nonewline; write-host -ForegroundColor green " is available. Test duration (Seconds): " -NoNewline;Write-Host $duration
                        }
                    }
                    
                    $reportfinding 
                    if($isAvailable){
                        $reportfinding.Success = $true
                    }
    
                    Start-Sleep -Milliseconds 150
    
                }
        }
         
    }
    $Findings | Export-csv "$env:USERPROFILE\Desktop\O365UrlConnectionStatusFromHost$($env:COMPUTERNAME).csv" -NoTypeInformation -Force
    
    Write-host -foregroundcolor DarkCyan "The output file was saved to: " -nonewline; "$env:USERPROFILE\Desktop\O365UrlConnectionStatusFromHost$($env:COMPUTERNAME).csv" 
    
}
