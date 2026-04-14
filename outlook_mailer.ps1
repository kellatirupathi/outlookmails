param(
    [Parameter(Mandatory = $true)]
    [ValidateSet("list-accounts", "deliver")]
    [string]$Action,

    [string]$PayloadPath
)

$ErrorActionPreference = "Stop"

function Write-Json([hashtable]$Value) {
    $Value | ConvertTo-Json -Depth 10 -Compress
}

function Get-OutlookApplication {
    New-Object -ComObject Outlook.Application
}

function Get-OutlookAccounts {
    try {
        $outlook = Get-OutlookApplication
        $session = $outlook.Session
        $accounts = @($session.Accounts | ForEach-Object {
                if (-not [string]::IsNullOrWhiteSpace($_.SmtpAddress)) {
                    [string]$_.SmtpAddress
                }
            })

        Write-Output (Write-Json @{
                success  = $true
                accounts = $accounts
            })
    }
    catch {
        Write-Output (Write-Json @{
                success  = $false
                accounts = @()
                error    = $_.Exception.Message
            })
        exit 1
    }
}

function Deliver-OutlookMessages {
    if (-not $PayloadPath) {
        Write-Output (Write-Json @{
                success = $false
                error   = "PayloadPath is required for deliver action."
            })
        exit 1
    }

    try {
        $payload = Get-Content -LiteralPath $PayloadPath -Raw | ConvertFrom-Json -Depth 10
        $outlook = Get-OutlookApplication
        $session = $outlook.Session

        $senderAccount = $null
        if (-not [string]::IsNullOrWhiteSpace([string]$payload.sender_account)) {
            $senderAccount = $session.Accounts | Where-Object { $_.SmtpAddress -eq [string]$payload.sender_account } | Select-Object -First 1
            if (-not $senderAccount) {
                throw "Outlook account not found: $($payload.sender_account)"
            }
        }

        $results = @()
        foreach ($message in $payload.messages) {
            $mail = $outlook.CreateItem(0)
            $mail.To = [string]$message.to
            $mail.Subject = [string]$message.subject

            if ($senderAccount) {
                $null = $mail.GetInspector()
                $mail.SendUsingAccount = $senderAccount
            }

            if ([string]$message.body_mode -eq "HTML") {
                $mail.HTMLBody = [string]$message.body
            }
            else {
                $mail.Body = [string]$message.body
            }

            foreach ($attachment in @($message.attachments)) {
                if (-not [string]::IsNullOrWhiteSpace([string]$attachment)) {
                    $resolved = Resolve-Path -LiteralPath ([string]$attachment)
                    $null = $mail.Attachments.Add($resolved.Path)
                }
            }

            if ([bool]$payload.draft_only) {
                $mail.Save()
                $status = "draft_saved"
            }
            else {
                $mail.Send()
                $status = "sent"
            }

            $results += @{
                to     = [string]$message.to
                status = $status
            }
        }

        Write-Output (Write-Json @{
                success   = $true
                processed = $results.Count
                results   = $results
            })
    }
    catch {
        Write-Output (Write-Json @{
                success = $false
                error   = $_.Exception.Message
            })
        exit 1
    }
}

switch ($Action) {
    "list-accounts" { Get-OutlookAccounts }
    "deliver" { Deliver-OutlookMessages }
}
