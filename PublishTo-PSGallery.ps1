# my "internal" policy is that code is never published to production unless it has been
# fully checked in to Git 
if (Get-GitOkToDeploy) {
    Publish-Module -Path "$PSScriptRoot\PSUnlock-OfficeProtection" -NuGetApiKey $env:NUGET_API_KEY
}