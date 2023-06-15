<#
.SYNOPSIS
Convert a Sharepoint Online Time zone ID to a human readable string.

.NOTES
By Andreas Dieckmann - https://diecknet.de
Timezone IDs according to https://docs.microsoft.com/en-us/dotnet/api/microsoft.sharepoint.spregionalsettings.TimeZones?view=sharepoint-server#Microsoft_SharePoint_SPRegionalSettings_TimeZones

Licensed under MIT License
Copyright 2021 Andreas Dieckmann

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

.EXAMPLE
Convert-SPOTimezoneToString 14
(UTC-09:00) Alaska

.LINK
https://diecknet.de/en/2021/07/09/Sharepoint-Online-Timezones-by-PowerShell/
#>

function Convert-SPOTimezoneToString
(
# ID of a SPO Timezone
[int]$ID
) 

{

    $TimeZoneIDs = @{
        39="(UTC-12:00) International Date Line West";
        95="(UTC-11:00) Coordinated Universal Time-11";
        15="(UTC-10:00) Hawaii";
        14="(UTC-09:00) Alaska";
        78="(UTC-08:00) Baja California";
        13="(UTC-08:00) Pacific Time (US and Canada)";
        38="(UTC-07:00) Arizona";
        77="(UTC-07:00) Chihuahua, La Paz, Mazatlan";
        12="(UTC-07:00) Mountain Time (US and Canada)";
        55="(UTC-06:00) Central America";
        11="(UTC-06:00) Central Time (US and Canada)";
        37="(UTC-06:00) Guadalajara, Mexico City, Monterrey";
        36="(UTC-06:00) Saskatchewan";
        35="(UTC-05:00) Bogota, Lima, Quito";
        10="(UTC-05:00) Eastern Time (US and Canada)";
        34="(UTC-05:00) Indiana (East)";
        88="(UTC-04:30) Caracas";
        91="(UTC-04:00) Asuncion";
        9="(UTC-04:00) Atlantic Time (Canada)";
        81="(UTC-04:00) Cuiaba";
        33="(UTC-04:00) Georgetown, La Paz, Manaus, San Juan";
        28="(UTC-03:30) Newfoundland";
        8="(UTC-03:00) Brasilia";
        85="(UTC-03:00) Buenos Aires";
        32="(UTC-03:00) Cayenne, Fortaleza";
        60="(UTC-03:00) Greenland";
        90="(UTC-03:00) Montevideo";
        103="(UTC-03:00) Salvador";
        65="(UTC-03:00) Santiago";
        96="(UTC-02:00) Coordinated Universal Time-02";
        30="(UTC-02:00) Mid-Atlantic";
        29="(UTC-01:00) Azores";
        53="(UTC-01:00) Cabo Verde";
        86="(UTC) Casablanca";
        93="(UTC) Coordinated Universal Time";
        2="(UTC) Dublin, Edinburgh, Lisbon, London";
        31="(UTC) Monrovia, Reykjavik";
        4="(UTC+01:00) Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna";
        6="(UTC+01:00) Belgrade, Bratislava, Budapest, Ljubljana, Prague";
        3="(UTC+01:00) Brussels, Copenhagen, Madrid, Paris";
        57="(UTC+01:00) Sarajevo, Skopje, Warsaw, Zagreb";
        69="(UTC+01:00) West Central Africa";
        83="(UTC+01:00) Windhoek";
        79="(UTC+02:00) Amman";
        5="(UTC+02:00) Athens, Bucharest, Istanbul";
        80="(UTC+02:00) Beirut";
        49="(UTC+02:00) Cairo";
        98="(UTC+02:00) Damascus";
        50="(UTC+02:00) Harare, Pretoria";
        59="(UTC+02:00) Helsinki, Kyiv, Riga, Sofia, Tallinn, Vilnius";
        101="(UTC+02:00) Istanbul";
        27="(UTC+02:00) Jerusalem";
        7="(UTC+02:00) Minsk (old)";
        104="(UTC+02:00) E. Europe";
        100="(UTC+02:00) Kaliningrad (RTZ 1)";
        26="(UTC+03:00) Baghdad";
        74="(UTC+03:00) Kuwait, Riyadh";
        109="(UTC+03:00) Minsk";
        51="(UTC+03:00) Moscow, St. Petersburg, Volgograd (RTZ 2)";
        56="(UTC+03:00) Nairobi";
        25="(UTC+03:30) Tehran";
        24="(UTC+04:00) Abu Dhabi, Muscat";
        54="(UTC+04:00) Baku";
        106="(UTC+04:00) Izhevsk, Samara (RTZ 3)";
        89="(UTC+04:00) Port Louis";
        82="(UTC+04:00) Tbilisi";
        84="(UTC+04:00) Yerevan";
        48="(UTC+04:30) Kabul";
        58="(UTC+05:00) Ekaterinburg (RTZ 4)";
        87="(UTC+05:00) Islamabad, Karachi";
        47="(UTC+05:00) Tashkent";
        23="(UTC+05:30) Chennai, Kolkata, Mumbai, New Delhi";
        66="(UTC+05:30) Sri Jayawardenepura";
        62="(UTC+05:45) Kathmandu";
        71="(UTC+06:00) Astana";
        102="(UTC+06:00) Dhaka";
        46="(UTC+06:00) Novosibirsk (RTZ 5)";
        61="(UTC+06:30) Yangon (Rangoon)";
        22="(UTC+07:00) Bangkok, Hanoi, Jakarta";
        64="(UTC+07:00) Krasnoyarsk (RTZ 6)";
        45="(UTC+08:00) Beijing, Chongqing, Hong Kong, Urumqi";
        63="(UTC+08:00) Irkutsk (RTZ 7)";
        21="(UTC+08:00) Kuala Lumpur, Singapore";
        73="(UTC+08:00) Perth";
        75="(UTC+08:00) Taipei";
        94="(UTC+08:00) Ulaanbaatar";
        20="(UTC+09:00) Osaka, Sapporo, Tokyo";
        72="(UTC+09:00) Seoul";
        70="(UTC+09:00) Yakutsk (RTZ 8)";
        19="(UTC+09:30) Adelaide";
        44="(UTC+09:30) Darwin";
        18="(UTC+10:00) Brisbane";
        76="(UTC+10:00) Canberra, Melbourne, Sydney";
        43="(UTC+10:00) Guam, Port Moresby";
        42="(UTC+10:00) Hobart";
        99="(UTC+10:00) Magadan";
        68="(UTC+10:00) Vladivostok, Magadan (RTZ 9)";
        107="(UTC+11:00) Chokurdakh (RTZ 10)";
        41="(UTC+11:00) Solomon Is., New Caledonia";
        108="(UTC+12:00) Anadyr, Petropavlovsk-Kamchatsky (RTZ 11)";
        17="(UTC+12:00) Auckland, Wellington";
        97="(UTC+12:00) Coordinated Universal Time+12";
        40="(UTC+12:00) Fiji";
        92="(UTC+12:00) Petropavlovsk-Kamchatsky - Old";
        67="(UTC+13:00) Nuku'alofa";
        16="(UTC+13:00) Samoa";
        }
        $TimeZoneString = $TimeZoneIDs.Get_Item($ID)
        if($null -ne $TimeZoneString) {
            return $TimeZoneString
        } else {
            return $ID
        }
    }   


<#
.SYNOPSIS
Get and return a CSV of all the Timezone and Regional settings for all (or some depending on the filters) of the SharePoint Online sites in an MS365 tenant.  

.NOTES
By Kent MacMillan - https://grump-it.pro

Be sure to enter your SPO admin site's URL as shown in the expample below.  Your normal tenant URL will not work for SPO commands.  Your account will of course have to have the appropriate read permissions to all SPO-sites in order to retrieve the time zone information.

Licensed under MIT License
Copyright 2023 Kent MacMillan

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

.EXAMPLE
$AdminSiteURL = "SPO ADMIN SITE URL"
Connect-SPOService -URL $AdminSiteURL
getSPOSitesTimeZones -csvExportPath "C:\Temp"

.LINK
https://grump-it.pro/blog/2023/06/15/change-all-time-zones-and-regional-settings-in-sharepoint-online/
#>

function getSPOSitesTimeZones

    [CmdletBinding()]
    param
    (
        [Parameter()]
        [string]$csvExportPath
    )

{
    $SPOSites = Get-SPOSite -Limit All | Where-Object {($_.URL -inotmatch 'appcatalog') -or ($_.URL -inotmatch '/portals/hub') -or ($_.URL -inotmatch 'my.sharepoint.com/')} 
    Foreach($SPOSite in $SPOSites)
    {
        try 
        {
            # Trying to retrieve regional settings of $SPOSite
            $regionalSettings = Get-SPOSiteScriptFromWeb -WebUrl $SPOSite.Url -IncludeRegionalSettings | ConvertFrom-Json
            $TimeZoneID = $regionalSettings.actions.timeZone
        } 
        catch 
        {
            # failback to 0 if not found
            $TimeZoneID = 0
        }

        $TimeZoneName = Convert-SPOTimezoneToString -ID $TimeZoneID

        Write-Host "SPO Site:" $SPOSite.URL " has the TimeZone of" $TimezoneName

        $CSVObject = [pscustomobject]@{
            'SPO URL' = $SPOSite.Url;
            'Zeitzone' = $TimezoneName;
            'Locale' = $regionalSettings.actions.locale
            'Zeitformat' = $regionalSettings.actions.hourFormat
            'Template' = $SPOSite.Template;
            'IsTeamsSite' = $SPOSite.IsTeamsConnected;
            'GroupID' = $SPOSite.RelatedGroupId
        }

        $Date = Get-Date -Format "dd-MM-yy HH_mm"
        $CSVPath = $csvExportPath + "\SPOSite_RegionaleSettings_Export_"+$Date+".csv" 

        $CSVObject | Export-CSV -Path $CSVPath -Encoding UTF8 -Delimiter ";" -NoClobber -Append -Force -NoTypeInformation
    }

}

<#
.SYNOPSIS
Create a "Site Design" (json script object) with the appropriate, new regional settings, including time zone, and apply these settings and the site design to all SPO sites which currently do not have them.

.NOTES
By Kent MacMillan - https://grump-it.pro

This script is currently set to create a "site design" as explained here: https://learn.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview.  This site design is set for German speaking tenant with Central European Time as the time zone.  It will check first if the site is set to proper de-DE regional code of 1031, and if not it applies the new site design/regional settings.  Furthermore the script will filter out non-standard SPO-sites, so that no settings are affected on the built-in, system sites such as the APPCATALOG#0 site, for example.  It does use the older CSOM method for applying these changes so be sure that your local machine from which you will execute the script meets the requirements for CSOM connections to SP: https://www.sharepointdiary.com/2019/04/connect-to-sharepoint-online-using-csom-powershell.html

Licensed under MIT License
Copyright 2023 Kent MacMillan

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

.EXAMPLE
$AdminSiteURL = "CHANGE ME"
Connect-SPOService -URL $AdminSiteURL

#Load SharePoint CSOM Assemblies
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

#Get Credentials to connect via CSOM
$Cred = Get-Credential
$Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Cred.Username, $Cred.Password)

setSPOSitesTimeZonesToDEStandard -adminUserName "CHANGE ME"

.LINK
https://grump-it.pro/blog/2023/06/15/change-all-time-zones-and-regional-settings-in-sharepoint-online/
#>

function setSPOSitesTimeZonesToDEStandard

    [CmdletBinding()]
    param
    (
        [Parameter()]
        [string]$adminUserName
    )
	
{
	
$script = @"
{
    "$schema": "schema.json",
        "actions": [
                        {
                            "verb": "setRegionalSettings",
                            "timeZone": 4, /* (UTC+01:00) Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna */
                            "locale": 1031, /* de-DE Format */
                            "sortOrder": 25, 
                            "hourFormat": "24" /* 24-Stunden-Uhr */
                        }
                   ],
    "bindata": { },
    "version": 1
}
"@

    $DE_StandardSiteDesign = Get-SPOSiteDesign | Where-Object -Property "Title" -EQ "DE_Standard"
    
    if(!$DE_StandardSiteDesign -or $DE_StandardSiteDesign -eq $null)
    {

        Write-Host "Standard DE site design not found...creating" -ForegroundColor Yellow
    
        Try
        {
            Add-SPOSiteScript -Title "DE_Standard" -Description "Standard regionale Einstellungen fuer DE" -Content $script -Verbose
            $siteScript = Get-SPOSiteScript | Where-Object -Property "Title" -EQ "DE_Standard"
            Add-SPOSiteDesign -Title "DE_Standard" -WebTemplate "64" -SiteScripts $siteScript.ID -Description "Standard regionale Einstellungen fuer DE" -IsDefault -Verbose
        }
        Catch
        {
             Write-Host "Standard DE site design could not be created...something went wrong" -ForegroundColor Red
             Write-Warning $Error[0]
        }
    }
    elseif($DE_StandardSiteDesign)
    {

        $SPOSites = Get-SPOSite -Limit All | Where-Object {(($_.URL -inotmatch 'appcatalog') -or ($_.URL -inotmatch '/portals/hub') -or ($_.URL -inotmatch 'my.sharepoint.com/'))}
    
        Foreach($SPOSite in $SPOSites)
        {
            try 
            {
            # trying to retrieve regional settings of $SPOSite
                $regionalSettings = Get-SPOSiteScriptFromWeb -WebUrl $SPOSite.Url -IncludeRegionalSettings | ConvertFrom-Json
                $currentTimeZoneID = $regionalSettings.actions.timeZone
                $currentLocale = $regionalSettings.actions.locale
            } 
            catch 
            {
                # failback to 0 if not found
                $currentTimeZoneID = 0
            }

            $currentTimeZoneName = Convert-SPOTimezoneToString -ID $currentTimeZoneID
            Write-Host "SPO Site:" $SPOSite.URL "has the current Time Zone of" $currentTimeZoneName -ForegroundColor Cyan
            Write-Host "and the locale of $currentLocale"  -ForegroundColor Cyan
            Write-Host "SPO Site: and is the type of" $SPOSite.Template -ForegroundColor Cyan
   
            If(($currentTimeZoneName -ne "(UTC+01:00) Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna" -or $currentLocale -ne "1031") -and $SPOSite.IsHubSite -eq $false -and `
                 ($SPOSite.Template -eq "EHS#1" -or $SPOSite.Template -eq "GROUP#0" -or $SPOSite.Template -eq "STS#3") -and ($SPOSite.Template -ne "APPCATALOG#0" -or ` 
                 $SPOSite.Template -ne "SRCHCEN#0" -or $SPOSite.Template -ne "SPSMSITEHOST#0" -or $SPOSite.Template -ne "RedirectSite#0" -or $SPOSite.Template -ne "PWA#0" -or ` 
                 $SPOSite.Template -ne "POINTPUBLISHINGTOPIC#0" -or $SPOSite.Template -ne "POINTPUBLISHINGHUB#0"))
            {
                Write-Host "Changing SPO Group or Communications Site:" $SPOSite.URL "to the time zone of (UTC+01:00) Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna" -ForegroundColor Yellow
                Write-Host "and the hour format to 24" -ForegroundColor Yellow
                Write-Host "and the regional format to de-DE" -ForegroundColor Yellow
                " " 
                Invoke-SPOSiteDesign -Identity $DE_StandardSiteDesign.ID -WebUrl $SPOSite.URL  -ErrorAction Stop -Verbose
            }
            Elseif(($currentTimeZoneName -ne "(UTC+01:00) Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna" -or $currentLocale -ne "1031") -and $SPOSite.IsHubSite -eq $false -and `
                ($SPOSite.Template -match "TEAMCHANNEL#0" -or $SPOSite.Template -match "TEAMCHANNEL#1"))
            {

                Try
                {
                    $adminAccount = Get-SPOUser -LoginName $adminUserName -Site $SPOSite.URL
                }
                Catch
                {
                    Write-Host "Admin Account is not a member of the site:" $SPOSite.URL  -ForegroundColor Red
                    Write-Warning $Error[0]
                }

                If(!$adminAccount -or $adminUserName.IsSiteAdmin -eq $False)
                {
                    Write-Host "Adding:" $adminUserName "to" $SPOSite.URL "as site collection admin" -ForegroundColor Green
                    Set-SPOUser -Site $SPOSite.URL -LoginName $adminUserName -IsSiteCollectionAdmin $True -Verbose
                }

                #Set up the context
                $Ctx = New-Object Microsoft.SharePoint.Client.ClientContext($SPOSite.URL)
                $Ctx.Credentials = $Credentials

                #Config parameters for SharePoint Online Site URL and Timezone description
                $TimezoneName ="(UTC+01:00) Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna"
                $LocaleName = "de-DE"
 
                #Get all available time zones
                $Timezones = $Ctx.Web.RegionalSettings.TimeZones
                $Ctx.Load($Timezones)
                $Ctx.ExecuteQuery()

                #$Locales = $Ctx.Web.RegionalSettings.LocaleID
                #"help"
                #$Locales
                #$Ctx.Load($Locales)
                #$Ctx.ExecuteQuery()
 
                #Filter the Time zone to update
                $NewTimezone = $Timezones | Where {$_.Description -eq $TimezoneName}
                #$NewLocale = $Locales | Where {$_.Description -eq $LocaleName}

                Write-Host "Changing SPO Team Site:" $SPOSite.URL "to the time zone of (UTC+01:00) Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna" -ForegroundColor Gray
                Write-Host "and the hour format to 24" -ForegroundColor Gray
                Write-Host "and the regional format to de-DE" -ForegroundColor Gray
                " "
                #sharepoint online powershell set time zone
                $Ctx.Web.RegionalSettings.TimeZone = $NewTimezone
                $Ctx.Web.RegionalSettings.LocaleID = "1031"
                $Ctx.Web.RegionalSettings.time24 = $True
                $Ctx.Web.Update()
                $Ctx.ExecuteQuery()

                Try
                {
                    Write-Host "Removing:" $adminUserName "to" $SPOSite.URL "as site collection admin" -ForegroundColor Green
                    Set-SPOUser -Site $SPOSite.URL -LoginName $adminUserName -IsSiteCollectionAdmin $False -Verbose
                    " "
                }
                Catch
                {
                    Write-Host "Admin Account was not a member of the site:" $SPOSite.URL "but it should have been.."  -ForegroundColor Red
                    Write-Warning $Error[0]
                }
            }
            Else
            {
                Write-Host "...which matches the desired time zone of (UTC+01:00) Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna, thus not changing" $SPOSite.URL -ForegroundColor Green
                " " 
            }
        
        }
    }
    Else
    {
        Write-Host "Standard DE site design not found...something went wrong" -ForegroundColor Red
        " "
    }
}

#Connect to SPO Service
#$AdminSiteURL = "CHANGE ME"
#Connect-SPOService -URL $AdminSiteURL

#Load SharePoint CSOM Assemblies
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

#Get Credentials to connect
#$Cred = Get-Credential
#$Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Cred.Username, $Cred.Password)

getSPOSitesTimeZones
#setSPOSitesTimeZonesToDEStandard -adminUserName "CHANGE ME"
