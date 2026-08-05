# 받은 편지함(및 하위 폴더)에서 조건에 맞는 메일을 검색합니다.
# 날짜/폴더 등의 로케일 문제를 피하기 위해 Outlook Restrict 대신
# 클라이언트 측(PowerShell)에서 직접 비교/필터링합니다.
#
# Request:  {
#   folderPath?: string, subject?: string, body?: string, sender?: string,
#   keywords?: string[], dateFrom?: string, dateTo?: string, maxCount?: number
# }
# Response: { items: MailSummary[], scanned: number, truncated: bool }
param(
    [Parameter(Mandatory)][string]$RequestPath,
    [Parameter(Mandatory)][string]$ResponsePath
)

. "$PSScriptRoot\_common.ps1"

function Search-FolderTree {
    param(
        $Folder,
        $InboxFolder,
        $Filters,
        [System.Collections.Generic.List[object]]$Results,
        [ref]$ScanCount,
        [int]$MaxCount,
        [int]$MaxScan
    )

    if ($Results.Count -ge $MaxCount -or $ScanCount.Value -ge $MaxScan) { return }

    $items = $Folder.Items
    try { $items.Sort('[ReceivedTime]', $true) } catch {}

    foreach ($item in $items) {
        if ($Results.Count -ge $MaxCount -or $ScanCount.Value -ge $MaxScan) { break }
        if (-not (Get-ItemType ([int]$item.Class))) { continue }
        $ScanCount.Value = $ScanCount.Value + 1

        if ($null -ne $Filters.dateFromDate -and $item.ReceivedTime -lt $Filters.dateFromDate) { continue }
        if ($null -ne $Filters.dateToDate -and $item.ReceivedTime -gt $Filters.dateToDate) { continue }

        if ($Filters.subject) {
            $subj = ''
            try { $subj = ([string]$item.Subject).ToLower() } catch {}
            if (-not $subj.Contains($Filters.subject)) { continue }
        }

        if ($Filters.sender) {
            $senderText = (([string]$item.SenderName) + ' ' + (Get-SenderSmtp $item)).ToLower()
            if (-not $senderText.Contains($Filters.sender)) { continue }
        }

        $needBody = [bool]$Filters.body -or ($Filters.keywords.Count -gt 0)
        $bodyText = ''
        if ($needBody) {
            try { $bodyText = [string]$item.Body } catch {}
        }

        if ($Filters.body -and -not ($bodyText.ToLower().Contains($Filters.body))) { continue }

        if ($Filters.keywords.Count -gt 0) {
            $cc = Get-CcInfo $item
            $haystack = (([string]$item.Subject) + ' ' + $bodyText + ' ' + ([string]$item.SenderName) + ' ' + ($cc.Names -join ' ') + ' ' + ($cc.Emails -join ' ')).ToLower()
            $matched = $false
            foreach ($kw in $Filters.keywords) {
                if ($haystack.Contains($kw.ToLower())) { $matched = $true; break }
            }
            if (-not $matched) { continue }
        }

        $Results.Add((ConvertTo-MailSummary $item $InboxFolder)) | Out-Null
    }

    foreach ($sub in $Folder.Folders) {
        if ($Results.Count -ge $MaxCount -or $ScanCount.Value -ge $MaxScan) { break }
        Search-FolderTree -Folder $sub -InboxFolder $InboxFolder -Filters $Filters -Results $Results -ScanCount $ScanCount -MaxCount $MaxCount -MaxScan $MaxScan
    }
}

try {
    $req = Read-Request $RequestPath

    $app = Get-OutlookApp
    $ns = Get-OutlookNamespace $app
    $inbox = Get-InboxFolder $ns

    $folderPathParam = Get-Prop $req 'folderPath' ''
    $root = Resolve-FolderPath -Root $inbox -RelativePath $folderPathParam -CreateIfMissing $false

    $filters = [ordered]@{
        subject      = $null
        sender       = $null
        body         = $null
        keywords     = @()
        dateFromDate = $null
        dateToDate   = $null
    }

    $subjectRaw = Get-Prop $req 'subject'
    if ($subjectRaw) { $filters.subject = ([string]$subjectRaw).ToLower() }

    $senderRaw = Get-Prop $req 'sender'
    if ($senderRaw) { $filters.sender = ([string]$senderRaw).ToLower() }

    $bodyRaw = Get-Prop $req 'body'
    if ($bodyRaw) { $filters.body = ([string]$bodyRaw).ToLower() }

    $kwRaw = Get-Prop $req 'keywords' @()
    if ($kwRaw) { $filters.keywords = @($kwRaw) }

    $dateFromRaw = Get-Prop $req 'dateFrom'
    if ($dateFromRaw) { $filters.dateFromDate = [DateTime]::Parse($dateFromRaw, [System.Globalization.CultureInfo]::InvariantCulture) }

    $dateToRaw = Get-Prop $req 'dateTo'
    if ($dateToRaw) { $filters.dateToDate = [DateTime]::Parse($dateToRaw, [System.Globalization.CultureInfo]::InvariantCulture) }

    $maxCount = [int](Get-Prop $req 'maxCount' 25)
    $maxScan = 5000

    $results = New-Object System.Collections.Generic.List[object]
    $scanCount = 0

    Search-FolderTree -Folder $root -InboxFolder $inbox -Filters $filters -Results $results -ScanCount ([ref]$scanCount) -MaxCount $maxCount -MaxScan $maxScan

    $data = [ordered]@{ items = $results; scanned = $scanCount; truncated = ($scanCount -ge $maxScan) }
    Write-JsonResult -ResponsePath $ResponsePath -Ok $true -Data $data
} catch {
    Write-JsonResult -ResponsePath $ResponsePath -Ok $false -ErrorMessage $_.Exception.Message
    exit 1
}
