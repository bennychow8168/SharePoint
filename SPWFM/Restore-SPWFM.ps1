# --- 1. Configuration Variables ---
$SQLInstance = "SQLServerName" # e.g., "DBServer\Instance"
$RunAsAccount = "DOMAIN\ServiceAccount"
$RunAsPassword = ConvertTo-SecureString "YourAccountPassword" -AsPlainText -Force
$CertGenKey = ConvertTo-SecureString "YourOriginalCertKey" -AsPlainText -Force

# Database Names (Update these if you used custom names)
$SBGatewayDB = "SbGatewayDatabase"
$SBContainerDB = "SBMessageContainer01"
$SBManagementDB = "SbManagementDB"
$WFInstanceDB = "WFInstanceManagementDB"
$WFResourceDB = "WFResourceManagementDB"
$WFManagementDB = "WFManagementDB"

# --- 2. Restore Service Bus Farm ---
Write-Host "Restoring Service Bus Farm..." -ForegroundColor Cyan

Restore-SBFarm `
    -RunAsAccount $RunAsAccount `
    -GatewayDBConnectionString "Data Source=$SQLInstance;Initial Catalog=$SBGatewayDB;Integrated Security=True" `
    -SBFarmDBConnectionString "Data Source=$SQLInstance;Initial Catalog=$SBManagementDB;Integrated Security=True" `
    -CertificateGenerationKey $CertGenKey

# --- 3. Add Local Host to Service Bus ---
Write-Host "Adding local host to Service Bus..." -ForegroundColor Cyan

Add-SBHost `
    -SBFarmDBConnectionString "Data Source=$SQLInstance;Initial Catalog=$SBManagementDB;Integrated Security=True" `
    -RunAsPassword $RunAsPassword `
    -CertificateGenerationKey $CertGenKey

# --- 4. Restore Workflow Manager Farm ---
Write-Host "Restoring Workflow Manager Farm..." -ForegroundColor Cyan

Restore-WFFarm `
    -RunAsAccount $RunAsAccount `
    -InstanceDBConnectionString "Data Source=$SQLInstance;Initial Catalog=$WFInstanceDB;Integrated Security=True" `
    -ResourceDBConnectionString "Data Source=$SQLInstance;Initial Catalog=$WFResourceDB;Integrated Security=True" `
    -MgmtDBConnectionString "Data Source=$SQLInstance;Initial Catalog=$WFManagementDB;Integrated Security=True" `
    -CertificateGenerationKey $CertGenKey

# --- 5. Add Local Host to Workflow Manager ---
Write-Host "Adding local host to Workflow Manager..." -ForegroundColor Cyan

Add-WFHost `
    -WFFarmDBConnectionString "Data Source=$SQLInstance;Initial Catalog=$WFManagementDB;Integrated Security=True" `
    -RunAsPassword $RunAsPassword `
    -CertificateGenerationKey $CertGenKey

Write-Host "SPWFM Restoration Complete. Please check services (Workflow Management & Service Bus Gateway) are running." -ForegroundColor Green
