# =============================================================================
# outlook-mcp 공통 PowerShell 헬퍼
# 모든 스크립트가 ". "$PSScriptRoot\_common.ps1"" 형태로 dot-source 하여 사용합니다.
# Windows PowerShell 5.1 (Classic Outlook COM) 기준으로 작성되었습니다.
# =============================================================================

$ErrorActionPreference = 'Stop'

# --- 상수 (Outlook Object Model) ---------------------------------------------
$script:olFolderInbox = 6
$script:olFolderDrafts = 16
$script:olFolderSentMail = 5
$script:olMail = 43          # OlObjectClass.olMail
$script:olTo = 1             # OlMailRecipientType.olTo
$script:olCC = 2             # OlMailRecipientType.olCC
$script:olMailItem = 0       # OlItemType.olMailItem

# 회의 요청/응답, 작업 요청 등 "캘린더성" 항목 클래스 (연차 신청, 회의 초대 등이
# 받은편지함에 이 클래스로 들어온다). organize_mail 은 이들을 calendar_folder 로 보낸다.
$script:olMeetingClasses = @(53, 54, 55, 56, 57, 49, 50, 51, 52)  # Meeting*, TaskRequest*

function Get-ItemType {
    param([int]$Class)
    if ($Class -eq $script:olMail) { return 'mail' }
    if ($script:olMeetingClasses -contains $Class) { return 'calendar' }
    return $null
}

# --- 요청/응답 JSON 처리 ------------------------------------------------------

function Read-Request {
    param([Parameter(Mandatory)][string]$Path)
    $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    if ([string]::IsNullOrWhiteSpace($raw)) {
        return New-Object PSObject
    }
    return $raw | ConvertFrom-Json
}

function Write-JsonResult {
    param(
        [Parameter(Mandatory)][string]$ResponsePath,
        [Parameter(Mandatory)][bool]$Ok,
        $Data = $null,
        [string]$ErrorMessage = $null
    )
    $payload = [ordered]@{ ok = $Ok }
    if ($Ok) {
        $payload.data = $Data
    } else {
        $payload.error = $ErrorMessage
    }
    $json = $payload | ConvertTo-Json -Depth 20 -Compress
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($ResponsePath, $json, $utf8NoBom)
}

function Get-Prop {
    param($Obj, [Parameter(Mandatory)][string]$Name, $Default = $null)
    if ($null -eq $Obj) { return $Default }
    $prop = $Obj.PSObject.Properties[$Name]
    if ($null -eq $prop -or $null -eq $prop.Value) { return $Default }
    return $prop.Value
}

# --- Outlook COM 연결 ---------------------------------------------------------

function Get-OutlookApp {
    try {
        # 이미 실행 중인 Outlook 인스턴스에 붙습니다 (새 창/프로필 프롬프트 방지).
        return [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
    } catch {
        return New-Object -ComObject Outlook.Application
    }
}

function Get-OutlookNamespace {
    param($App)
    return $App.GetNamespace('MAPI')
}

function Get-InboxFolder {
    param($Ns)
    return $Ns.GetDefaultFolder($script:olFolderInbox)
}

function Get-DraftsFolder {
    param($Ns)
    return $Ns.GetDefaultFolder($script:olFolderDrafts)
}

function Get-SentFolder {
    param($Ns)
    return $Ns.GetDefaultFolder($script:olFolderSentMail)
}

# organize_mail 등이 받은편지함/보낸편지함 중 어디를 기준으로 동작할지 선택할 때 사용.
function Get-RootFolder {
    param($Ns, [string]$RootFolderName)
    if ($RootFolderName -eq 'sent') { return Get-SentFolder $Ns }
    return Get-InboxFolder $Ns
}

function Get-ItemByIds {
    param($Ns, [Parameter(Mandatory)][string]$EntryId, [string]$StoreId)
    if ($StoreId) {
        return $Ns.GetItemFromID($EntryId, $StoreId)
    }
    return $Ns.GetItemFromID($EntryId)
}

# --- 폴더 경로 처리 ------------------------------------------------------------

function Resolve-FolderPath {
    param(
        $Root,
        [string]$RelativePath,
        [bool]$CreateIfMissing = $false
    )
    $current = $Root
    if ([string]::IsNullOrWhiteSpace($RelativePath)) {
        return $current
    }
    $segments = $RelativePath -split '/' | Where-Object { $_ -ne '' }
    foreach ($seg in $segments) {
        $next = $null
        foreach ($f in $current.Folders) {
            if ($f.Name -eq $seg) { $next = $f; break }
        }
        if ($null -eq $next) {
            if ($CreateIfMissing) {
                $next = $current.Folders.Add($seg)
            } else {
                throw "폴더를 찾을 수 없습니다: '$RelativePath' (없는 구간: '$seg')"
            }
        }
        $current = $next
    }
    return $current
}

function Get-RelativeFolderPath {
    param($Folder, $InboxFolder)
    try {
        $names = New-Object System.Collections.Generic.List[string]
        $cur = $Folder
        $depth = 0
        while ($null -ne $cur -and $cur.EntryID -ne $InboxFolder.EntryID -and $depth -lt 20) {
            $names.Insert(0, $cur.Name)
            $cur = $cur.Parent
            $depth++
        }
        return ($names -join '/')
    } catch {
        return ''
    }
}

# --- 메일 아이템 -> JSON 직렬화 헬퍼 -------------------------------------------

function Get-SenderSmtp {
    param($Item)
    try {
        $sender = $Item.Sender
        if ($null -ne $sender) {
            $exUser = $sender.GetExchangeUser()
            if ($null -ne $exUser) { return $exUser.PrimarySmtpAddress }
        }
    } catch {}
    try { return [string]$Item.SenderEmailAddress } catch { return '' }
}

function Get-RecipientSmtp {
    param($Recipient)
    try {
        $exUser = $Recipient.AddressEntry.GetExchangeUser()
        if ($null -ne $exUser) { return $exUser.PrimarySmtpAddress }
    } catch {}
    try { return [string]$Recipient.Address } catch { return '' }
}

function Get-RecipientsByType {
    param($Item, [int]$Type)
    $result = New-Object System.Collections.Generic.List[object]
    try {
        foreach ($r in $Item.Recipients) {
            if ($r.Type -eq $Type) {
                $result.Add([ordered]@{ name = $r.Name; email = (Get-RecipientSmtp $r) }) | Out-Null
            }
        }
    } catch {}
    # 주의: PowerShell 은 함수가 반환하는 컬렉션의 원소가 0개/1개면 자동으로 배열을 풀어버린다
    # (2개 이상일 때만 배열로 유지됨). 단항 콤마 연산자로 강제로 배열 형태를 유지시킨다.
    return ,$result
}

function Get-CcInfo {
    param($Item)
    $cc = Get-RecipientsByType $Item $script:olCC
    $names = @($cc | ForEach-Object { $_.name })
    $emails = @($cc | ForEach-Object { $_.email })
    return [ordered]@{ Names = $names; Emails = $emails }
}

function ConvertTo-MailSummary {
    param($Item, $InboxFolder)

    $body = ''
    try { $body = [string]$Item.Body } catch {}
    if ($body.Length -gt 2000) { $body = $body.Substring(0, 2000) }

    $folderPath = ''
    try { $folderPath = Get-RelativeFolderPath $Item.Parent $InboxFolder } catch {}

    $storeId = ''
    try { $storeId = $Item.Parent.StoreID } catch {}

    $subject = ''
    try { $subject = [string]$Item.Subject } catch {}

    $senderName = ''
    try { $senderName = [string]$Item.SenderName } catch {}

    $receivedTime = ''
    try { $receivedTime = $Item.ReceivedTime.ToString('o') } catch { $receivedTime = (Get-Date).ToString('o') }

    $unread = $false
    try { $unread = [bool]$Item.UnRead } catch {}

    $hasAttachments = $false
    $attachmentNames = @()
    try {
        $hasAttachments = ($Item.Attachments.Count -gt 0)
        $attachmentNames = @($Item.Attachments | ForEach-Object { $_.FileName })
    } catch {}

    $cc = Get-CcInfo $Item
    $itemType = Get-ItemType ([int]$Item.Class)

    return [ordered]@{
        entryId        = $Item.EntryID
        storeId        = $storeId
        subject        = $subject
        senderName     = $senderName
        senderEmail    = (Get-SenderSmtp $Item)
        receivedTime   = $receivedTime
        unread         = $unread
        hasAttachments = $hasAttachments
        attachmentNames = $attachmentNames
        bodyPreview    = $body
        ccNames        = $cc.Names
        ccEmails       = $cc.Emails
        folderPath     = $folderPath
        itemType       = if ($itemType) { $itemType } else { 'mail' }
    }
}
