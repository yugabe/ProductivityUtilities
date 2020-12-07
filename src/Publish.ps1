param (
    [string]$Project,
    [string]$Configuration = "Release",
    [string]$RuntimeIdentifier = "win-x64",
    [string]$PublishProtocol = "FileSystem",
    [bool]$SelfContained = $false,
    [bool]$PublishSingleFile = $true,
    [bool]$PublishReadyToRun = $true,
    [string]$PublishDir = $env:ProductivityUtilitiesInstallPath
)

if ($PublishDir -eq ""){
    $PublishDir = "$($env:APPDATA)/ProductivityUtilities"
}

[string[]]$vars = (Get-Command -Name $MyInvocation.InvocationName).Parameters | Select-Object -ExpandProperty Values | Select-Object -ExpandProperty Name | ForEach-Object -Process {Write-Host "$_ $($(Get-Variable -Name $_ -ErrorAction SilentlyContinue).Value)"};

dotnet publish $Project -c $Configuration -r $RuntimeIdentifier /p:PublishDir="$($PublishDir)" /p:PublishProtocol="$($PublishProtocol)" /p:SelfContained="$($SelfContained)" /p:PublishSingleFile="$($PublishSingleFile)" /p:PublishReadyToRun="$($PublishReadyToRun)"
