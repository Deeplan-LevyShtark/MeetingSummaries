# Define the path to your .env file
$envFilePath = ".\.env"

# Initialize an empty string for the UserName
$AppCatalogURL = ""
$UserName = ""
$Password = ""
Write-Host "Checking if the .env file exists at $envFilePath"

# Check if the .env file exists
if (Test-Path $envFilePath) {
    Write-Host ".env file found. Reading..."
    # Read the .env file and look for the USERNAME variable
    Get-Content $envFilePath | ForEach-Object {
        if ($_ -match '^USERNAME=') {
            # Extract the value after the equal sign
            $UserName = $_.Split('=')[1].Trim()
        }
        if ($_ -match '^PASSWORD=') {
            # Extract the value after the equal sign
            $Password = $_.Split('=')[1].Trim()
        }
        if ($_ -match '^TENANT=') {
            $AppCatalogURL = $_.Split('=')[1].Trim()
        }
    }
}
else {
    Write-Host ".env file does not exist at the specified path."
}

$SecurePassword = ConvertTo-SecureString -String $Password -AsPlainText -Force
$Cred = New-Object -TypeName System.Management.Automation.PSCredential -argumentlist $UserName, $SecurePassword

$AppFilePath = Get-ChildItem -Path "./sharepoint/solution/*.sppkg"

# Connect to SharePoint Online App Catalog site
try {
    Connect-PnPOnline -Url $AppCatalogURL -WarningAction Ignore -Credentials $Cred
}
catch {
    Write-Host "Error connecting to SharePoint Online: $_"
    exit
}

$content = Get-ChildItem -Path "./src"

# Webpart - Deploy to all sites, Extention - Deploy to one site
if ([string]$content[0] -eq "webparts") {
    Add-PnPApp -Path $AppFilePath -Scope Tenant -Overwrite -Publish -SkipFeatureDeployment
}
else {
    Add-PnPApp -Path $AppFilePath -Scope Tenant -Overwrite -Publish
}