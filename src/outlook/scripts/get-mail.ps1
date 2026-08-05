# 메일 한 건의 상세 정보를 가져옵니다.
# Request:  { entryId: string, storeId?: string }
# Response: MailDetail (MailSummary + to, cc, body, sentOn, attachments)
param(
    [Parameter(Mandatory)][string]$RequestPath,
    [Parameter(Mandatory)][string]$ResponsePath
)

. "$PSScriptRoot\_common.ps1"

try {
    $req = Read-Request $RequestPath
    $entryId = Get-Prop $req 'entryId'
    $storeId = Get-Prop $req 'storeId'
    if (-not $entryId) { throw 'entryId 파라미터가 필요합니다.' }

    $app = Get-OutlookApp
    $ns = Get-OutlookNamespace $app
    $inbox = Get-InboxFolder $ns

    $item = Get-ItemByIds $ns $entryId $storeId
    if ($item.Class -ne $script:olMail) {
        throw '메일(Mail) 항목이 아닙니다.'
    }

    $summary = ConvertTo-MailSummary $item $inbox

    $to = Get-RecipientsByType $item $script:olTo
    $cc = Get-RecipientsByType $item $script:olCC

    $body = ''
    try { $body = [string]$item.Body } catch {}

    $sentOn = $null
    try { $sentOn = $item.SentOn.ToString('o') } catch {}

    $attachments = New-Object System.Collections.Generic.List[object]
    foreach ($a in $item.Attachments) {
        $attachments.Add([ordered]@{ fileName = $a.FileName; size = $a.Size }) | Out-Null
    }

    $data = [ordered]@{
        entryId        = $summary.entryId
        storeId        = $summary.storeId
        subject        = $summary.subject
        senderName     = $summary.senderName
        senderEmail    = $summary.senderEmail
        receivedTime   = $summary.receivedTime
        sentOn         = $sentOn
        unread         = $summary.unread
        hasAttachments = $summary.hasAttachments
        bodyPreview    = $summary.bodyPreview
        ccNames        = $summary.ccNames
        ccEmails       = $summary.ccEmails
        folderPath     = $summary.folderPath
        to             = $to
        cc             = $cc
        body           = $body
        attachments    = $attachments
    }

    Write-JsonResult -ResponsePath $ResponsePath -Ok $true -Data $data
} catch {
    Write-JsonResult -ResponsePath $ResponsePath -Ok $false -ErrorMessage $_.Exception.Message
    exit 1
}
