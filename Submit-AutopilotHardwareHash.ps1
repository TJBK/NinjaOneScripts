$ErrorActionPreference = 'Stop'
$params = @{
    clientId = Ninja-Property-Get "autopilotazureappid"
    tenantId = Ninja-Property-Get "autopilotazuretenantid"
    secret   = Ninja-Property-Get "autopilotazureappsecret"
}
if ($params.Values -contains $null) {
    Ninja-Property-Set "autopilotuploadstatus" "<table border='1'><tr><td>Status</td><td>ERROR: Missing required parameters</td></tr></table>"
    exit 1
}
try {
    $token = (Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$($params.tenantId)/oauth2/v2.0/token" -Body @{
        client_id     = $params.clientId
        scope         = 'https://graph.microsoft.com/.default'
        client_secret = $params.secret
        grant_type    = 'client_credentials'
    }).access_token
    $serial = (Get-CimInstance -Class Win32_BIOS).SerialNumber
    $hardware = (Get-CimInstance -Namespace 'root\cimv2\mdm\dmmap' -Class 'MDM_DevDetail_Ext01' -Filter "InstanceID='Ext' AND ParentID='./DevDetail'").DeviceHardwareData
    $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem
    $deviceInfo = @"
<table border='1' style='border-collapse: collapse; width: 100%;'>
    <tr style='background-color: #f2f2f2;'><th>Property</th><th>Value</th></tr>
    <tr><td>Serial Number</td><td>$serial</td></tr>
    <tr><td>Manufacturer</td><td>$($computerSystem.Manufacturer)</td></tr>
    <tr><td>Model</td><td>$($computerSystem.Model)</td></tr>
"@
    $filter = [System.Web.HttpUtility]::UrlEncode("serialNumber eq '$serial'")
    $headers = @{
        "Authorization" = "Bearer $token"
        "Content-Type"  = "application/json"
    }
    $existing = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/deviceManagement/importedWindowsAutopilotDeviceIdentities?`$filter=$filter" -Headers $headers -Method Get
    if ($existing.value.Count -gt 0) {
        $deviceInfo += "<tr><td>Status</td><td>Device already exists in Autopilot</td></tr></table>"
        Ninja-Property-Set "AutopilotUploadStatus" $deviceInfo
        exit 0
    }
    $response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/deviceManagement/importedWindowsAutopilotDeviceIdentities" -Headers $headers -Method Post -Body (@{
        "@odata.type"               = "#microsoft.graph.importedWindowsAutopilotDeviceIdentity"
        "groupTag"                  = ""
        "serialNumber"              = $serial
        "hardwareIdentifier"        = $hardware
        "assignedUserPrincipalName" = ""
    } | ConvertTo-Json)
    
    1..30 | ForEach-Object {
        Start-Sleep -Seconds 10
        $status = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/deviceManagement/importedWindowsAutopilotDeviceIdentities?`$filter=$filter" -Headers $headers -Method Get
        if ($status.value.Count -gt 0) {
            $state = $status.value[0].state.deviceImportStatus
            switch ($state) {
                'complete' {
                    $deviceInfo += "<tr><td>Status</td><td>Successfully imported to Autopilot</td></tr></table>"
                    Ninja-Property-Set "autopilotuploadstatus" $deviceInfo
                    exit 0
                }
                'error' {
                    throw "Import failed: $($status.value[0].state.deviceErrorName) (Code: $($status.value[0].state.deviceErrorCode))"
                }
                'unknown' {
                    continue
                }
                'pending' {
                    continue
                }
                'partial' {
                    continue
                }
                default {
                    continue
                }
            }
        }
    }
    throw "Import timed out after 5 minutes"
}
catch {
    $deviceInfo += "<tr><td>Status</td><td>ERROR: $($_.Exception.Message)</td></tr></table>"
    Ninja-Property-Set "autopilotuploadstatus" $deviceInfo
    exit 1
}