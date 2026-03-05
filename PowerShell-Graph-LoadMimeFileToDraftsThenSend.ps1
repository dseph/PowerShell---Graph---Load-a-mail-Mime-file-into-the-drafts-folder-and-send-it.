# PowerShell-Graph-LoadMimeFileToDraftsThenSend.ps1
# Generated initially with CoPilot.
#
# =====================================================================================
# Load a MIME (.eml) file -> Create Draft in Drafts -> Send Draft
# Auth: Application (client credentials)
# Draft creation: POST /users/{id}/messages  (Content-Type: text/plain, body = base64 MIME)
# Send draft:     POST /users/{id}/messages/{id}/send (no body)
#
# Permissions needed:
# - Mail.ReadWrite (application) to create drafts in a user's mailbox. [1](https://learn.microsoft.com/en-us/graph/api/user-post-messages?view=graph-rest-1.0)
# - Mail.Send (application) to send messages from a user's mailbox. [3](https://learn.microsoft.com/en-us/graph/api/message-send?view=graph-rest-1.0)
# Don't forget to grant admin consent for the app permissions in Azure AD.
#
# Docs:
# - Create message (draft) supports MIME base64 with Content-Type: text/plain, saved to Drafts by default. [1](https://learn.microsoft.com/en-us/graph/api/user-post-messages?view=graph-rest-1.0)
# - Send an existing draft message: POST /users/{id}/messages/{id}/send. [3](https://learn.microsoft.com/en-us/graph/api/message-send?view=graph-rest-1.0)
# - MIME guidance and headers: [2](https://learn.microsoft.com/en-us/graph/outlook-send-mime-message)
# =====================================================================================

$ErrorActionPreference = "Stop"

# -----------------------
# CONFIG (EDIT THESE)
# -----------------------
$TenantId     = "YOUR_TENANT_ID"
$ClientId     = "YOUR_CLIENT_ID"
$ClientSecret = "YOUR_CLIENT_SECRET"
$UserId       = "user@contoso.com"           # mailbox to create/send from
$MimePath     = "C:\test\message.eml"       # path to RFC822/MIME file
# -----------------------

function Get-GraphToken {
    param(
        [Parameter(Mandatory)] [string] $TenantId,
        [Parameter(Mandatory)] [string] $ClientId,
        [Parameter(Mandatory)] [string] $ClientSecret
    )

    $tokenEndpoint = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $body = @{
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = "https://graph.microsoft.com/.default"
        grant_type    = "client_credentials"
    }

    $tokenResponse = Invoke-RestMethod -Method POST -Uri $tokenEndpoint -Body $body -ContentType "application/x-www-form-urlencoded"
    return $tokenResponse.access_token
}

function Invoke-Graph {
    param(
        [Parameter(Mandatory)] [ValidateSet("GET","POST","PATCH","DELETE")] [string] $Method,
        [Parameter(Mandatory)] [string] $Uri,
        [Parameter(Mandatory)] [string] $AccessToken,
        [string] $ContentType = "application/json",
        [object] $Body = $null
    )

    $headers = @{
        Authorization = "Bearer $AccessToken"
    }

    if ($null -ne $Body) {
        # If Body is already a string (e.g., base64 MIME), send it as-is.
        if ($Body -is [string]) {
            return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers -ContentType $ContentType -Body $Body
        }
        else {
            $json = $Body | ConvertTo-Json -Depth 20
            return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers -ContentType $ContentType -Body $json
        }
    }
    else {
        return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers -ContentType $ContentType
    }
}

# 1) Read MIME file bytes
if (-not (Test-Path $MimePath)) {
    throw "MIME file not found: $MimePath"
}
$mimeBytes = [System.IO.File]::ReadAllBytes($MimePath)

# 2) Base64 encode (Graph expects MIME content encoded in base64 in request body for MIME format) [1](https://learn.microsoft.com/en-us/graph/api/user-post-messages?view=graph-rest-1.0)[2](https://learn.microsoft.com/en-us/graph/outlook-send-mime-message)
$mimeBase64 = [System.Convert]::ToBase64String($mimeBytes)

# 3) Get token (client credentials)
$token = Get-GraphToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret

# 4) Create draft in Drafts (by default saved in Drafts folder) [1](https://learn.microsoft.com/en-us/graph/api/user-post-messages?view=graph-rest-1.0)
#    Must set Content-Type: text/plain when sending MIME format. [1](https://learn.microsoft.com/en-us/graph/api/user-post-messages?view=graph-rest-1.0)[2](https://learn.microsoft.com/en-us/graph/outlook-send-mime-message)
$createDraftUri = "https://graph.microsoft.com/v1.0/users/$UserId/messages"
$draft = Invoke-Graph -Method POST -Uri $createDraftUri -AccessToken $token -ContentType "text/plain" -Body $mimeBase64

Write-Host "Created draft."
Write-Host "Draft Id: $($draft.id)"
Write-Host "Subject : $($draft.subject)"

# 5) Send the draft (no body required) [3](https://learn.microsoft.com/en-us/graph/api/message-send?view=graph-rest-1.0)
$sendDraftUri = "https://graph.microsoft.com/v1.0/users/$UserId/messages/$($draft.id)/send"
Invoke-Graph -Method POST -Uri $sendDraftUri -AccessToken $token -ContentType "application/json"

Write-Host "Send requested (202 Accepted expected). Draft should now be in Sent Items."  # behavior described in send API [3](https://learn.microsoft.com/en-us/graph/api/message-send?view=graph-rest-1.0)
