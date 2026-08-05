# 답장/전체답장/새 메일 초안을 만들어 Outlook 임시보관함(Drafts)에 저장합니다.
# 발송은 하지 않습니다. mode 가 reply/replyAll 이면 Outlook 의 Reply()/ReplyAll() 을
# 사용하므로 원본 스레드 인용, 제목의 "RE:" 접두어, 수신자가 자동으로 채워집니다.
#
# Request:  {
#   mode: 'reply' | 'replyAll' | 'new',
#   sourceEntryId?: string, sourceStoreId?: string,
#   to?: string[], cc?: string[], subject?: string, bodyHtml: string
# }
# Response: { draftEntryId: string, draftStoreId: string, subject: string, to: string }
param(
    [Parameter(Mandatory)][string]$RequestPath,
    [Parameter(Mandatory)][string]$ResponsePath
)

. "$PSScriptRoot\_common.ps1"

try {
    $req = Read-Request $RequestPath
    $mode = Get-Prop $req 'mode'
    $bodyHtml = Get-Prop $req 'bodyHtml'
    if (-not $mode) { throw 'mode 파라미터가 필요합니다.' }
    if (-not $bodyHtml) { throw 'bodyHtml 파라미터가 필요합니다.' }

    $toList = Get-Prop $req 'to' @()
    $ccList = Get-Prop $req 'cc' @()
    $subject = Get-Prop $req 'subject'

    $app = Get-OutlookApp
    $ns = Get-OutlookNamespace $app

    if ($mode -eq 'new') {
        $mail = $app.CreateItem($script:olMailItem)
        if ($toList -and @($toList).Count -gt 0) { $mail.To = (@($toList) -join ';') }
        if ($ccList -and @($ccList).Count -gt 0) { $mail.CC = (@($ccList) -join ';') }
        if ($subject) { $mail.Subject = $subject }
        $mail.HTMLBody = $bodyHtml
    } else {
        $sourceEntryId = Get-Prop $req 'sourceEntryId'
        $sourceStoreId = Get-Prop $req 'sourceStoreId'
        if (-not $sourceEntryId) { throw 'reply/replyAll 모드에서는 sourceEntryId 파라미터가 필요합니다.' }

        $source = Get-ItemByIds $ns $sourceEntryId $sourceStoreId
        if ($mode -eq 'replyAll') {
            $mail = $source.ReplyAll()
        } elseif ($mode -eq 'reply') {
            $mail = $source.Reply()
        } else {
            throw "알 수 없는 mode 입니다: $mode"
        }

        if ($subject) { $mail.Subject = $subject }
        if ($toList -and @($toList).Count -gt 0) { $mail.To = (@($toList) -join ';') }
        if ($ccList -and @($ccList).Count -gt 0) { $mail.CC = (@($ccList) -join ';') }

        # 작성한 답장 본문을 원본 인용문 위에 삽입합니다.
        $mail.HTMLBody = $bodyHtml + $mail.HTMLBody
    }

    $mail.Save()

    $draftStoreId = ''
    try { $draftStoreId = $mail.Parent.StoreID } catch {}

    $data = [ordered]@{
        draftEntryId = $mail.EntryID
        draftStoreId = $draftStoreId
        subject      = [string]$mail.Subject
        to           = [string]$mail.To
    }

    Write-JsonResult -ResponsePath $ResponsePath -Ok $true -Data $data
} catch {
    Write-JsonResult -ResponsePath $ResponsePath -Ok $false -ErrorMessage $_.Exception.Message
    exit 1
}
