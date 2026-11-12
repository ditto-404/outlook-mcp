# 답장/전체답장/새 메일 초안을 만들어 Outlook 임시보관함(Drafts)에 저장합니다.
# 발송은 하지 않습니다. mode 가 reply/replyAll 이면 Outlook 의 Reply()/ReplyAll() 을
# 사용하므로 원본 스레드 인용, 제목의 "RE:" 접두어, 수신자가 자동으로 채워집니다.
# mode 가 update 이면 list-drafts.ps1 등으로 조회한 기존 초안(draftEntryId)을 새 초안으로
# 중복 생성하지 않고 그 자리에서 덮어씁니다(본문 전체 교체).
#
# Request:  {
#   mode: 'reply' | 'replyAll' | 'new' | 'update',
#   sourceEntryId?: string, sourceStoreId?: string,
#   draftEntryId?: string, draftStoreId?: string,
#   to?: string[], cc?: string[], subject?: string, bodyHtml: string,
#   includeSignature?: boolean (기본 true: 보낸편지함 최근 메일의 서명을 덧붙임)
# }
# Response: { draftEntryId: string, draftStoreId: string, subject: string, to: string }
#
# 본문은 보낸편지함 최근 메일에서 추출한 글꼴 스타일로 감싸 저장되므로,
# 초안이 사용자가 평소 보내는 메일과 같은 글꼴/서명으로 보입니다.
param(
    [Parameter(Mandatory)][string]$RequestPath,
    [Parameter(Mandatory)][string]$ResponsePath
)

. "$PSScriptRoot\_common.ps1"

# 보낸편지함에서 서식을 찾지 못했을 때 쓰는 기본 글꼴 (New Outlook/OWA 기본값).
$script:DefaultFontStyle = 'font-family: Aptos, Aptos_EmbeddedFont, Aptos_MSFontService, Calibri, Helvetica, sans-serif; font-size: 11pt; color: rgb(0, 0, 0);'

# 사용자가 실제로 보낸 최근 메일에서 본문 글꼴 스타일과 서명 HTML 을 추출합니다.
function Get-SentMailStyle {
    param($Ns)
    $result = @{ FontStyle = $null; SignatureHtml = '' }
    try {
        $items = $Ns.GetDefaultFolder($script:olFolderSentMail).Items
        $items.Sort('[SentOn]', $true)
        $checked = 0
        foreach ($m in $items) {
            if ($checked -ge 30) { break }
            $checked++
            $h = $null
            try {
                if ([int]$m.Class -ne $script:olMail) { continue }
                $h = [string]$m.HTMLBody
            } catch { continue }
            if (-not $h) { continue }
            $bodyStart = $h.IndexOf('<body')
            if ($bodyStart -lt 0) { continue }

            if (-not $result.FontStyle) {
                $mt = [regex]::Match($h.Substring($bodyStart), '<div[^>]*style="([^"]*font-family[^"]*)"')
                if ($mt.Success) { $result.FontStyle = $mt.Groups[1].Value }
            }
            if (-not $result.SignatureHtml) {
                $s = $h.IndexOf('<div id="Signature"')
                if ($s -ge 0) {
                    $ends = @('<div id="appendonsend"', '<hr', '<div id="divRplyFwdMsg"', '</body>') |
                        ForEach-Object { $h.IndexOf($_, $s) } | Where-Object { $_ -gt $s }
                    if ($ends) {
                        $e = [int](($ends | Measure-Object -Minimum).Minimum)
                        $result.SignatureHtml = $h.Substring($s, $e - $s)
                    }
                }
            }
            if ($result.FontStyle -and $result.SignatureHtml) { break }
        }
    } catch {}
    if (-not $result.FontStyle) { $result.FontStyle = $script:DefaultFontStyle }
    return $result
}

# 작성된 본문 HTML 을 사용자 글꼴로 감싸고, 필요하면 서명을 덧붙입니다.
function Build-StyledBody {
    param([string]$BodyHtml, $Style, [bool]$IncludeSignature)
    # <p> 기본 여백이 생기지 않도록 (사용자 메일은 줄 사이 여백 없음)
    $inner = $BodyHtml -replace '<p>', '<p style="margin:0">'
    $html = '<div style="' + $Style.FontStyle + '">' + $inner + '</div>'
    if ($IncludeSignature -and $Style.SignatureHtml) {
        $html += '<div style="' + $Style.FontStyle + '"><br></div>' + $Style.SignatureHtml
    }
    return $html
}

function New-HtmlDocument {
    param([string]$Inner)
    return '<html><head><meta http-equiv="Content-Type" content="text/html; charset=utf-8">' +
        '<style type="text/css" style="display:none;"> P {margin-top:0;margin-bottom:0;} </style>' +
        '</head><body dir="ltr">' + $Inner + '</body></html>'
}

try {
    $req = Read-Request $RequestPath
    $mode = Get-Prop $req 'mode'
    $bodyHtml = Get-Prop $req 'bodyHtml'
    if (-not $mode) { throw 'mode 파라미터가 필요합니다.' }
    if (-not $bodyHtml) { throw 'bodyHtml 파라미터가 필요합니다.' }

    $toList = Get-Prop $req 'to' @()
    $ccList = Get-Prop $req 'cc' @()
    $subject = Get-Prop $req 'subject'
    $includeSignature = [bool](Get-Prop $req 'includeSignature' $true)

    $app = Get-OutlookApp
    $ns = Get-OutlookNamespace $app

    $style = Get-SentMailStyle $ns
    $styledBody = Build-StyledBody -BodyHtml $bodyHtml -Style $style -IncludeSignature $includeSignature

    if ($mode -eq 'new') {
        $mail = $app.CreateItem($script:olMailItem)
        if ($toList -and @($toList).Count -gt 0) { $mail.To = (@($toList) -join ';') }
        if ($ccList -and @($ccList).Count -gt 0) { $mail.CC = (@($ccList) -join ';') }
        if ($subject) { $mail.Subject = $subject }
        $mail.HTMLBody = New-HtmlDocument $styledBody
    } elseif ($mode -eq 'update') {
        $draftEntryId = Get-Prop $req 'draftEntryId'
        $draftStoreId = Get-Prop $req 'draftStoreId'
        if (-not $draftEntryId) { throw 'update 모드에서는 draftEntryId 파라미터가 필요합니다.' }

        $mail = Get-ItemByIds $ns $draftEntryId $draftStoreId
        if ([int]$mail.Class -ne $script:olMail) { throw '메일(Mail) 항목이 아닙니다.' }

        if ($subject) { $mail.Subject = $subject }
        if ($toList -and @($toList).Count -gt 0) { $mail.To = (@($toList) -join ';') }
        if ($ccList -and @($ccList).Count -gt 0) { $mail.CC = (@($ccList) -join ';') }

        # update 모드는 본문 전체를 교체합니다 (reply 와 달리 원본 인용문을 덧붙이지 않음).
        $mail.HTMLBody = New-HtmlDocument $styledBody
    } elseif ($mode -eq 'reply' -or $mode -eq 'replyAll') {
        $sourceEntryId = Get-Prop $req 'sourceEntryId'
        $sourceStoreId = Get-Prop $req 'sourceStoreId'
        if (-not $sourceEntryId) { throw 'reply/replyAll 모드에서는 sourceEntryId 파라미터가 필요합니다.' }

        $source = Get-ItemByIds $ns $sourceEntryId $sourceStoreId
        if ($mode -eq 'replyAll') {
            $mail = $source.ReplyAll()
        } else {
            $mail = $source.Reply()
        }

        if ($subject) { $mail.Subject = $subject }
        if ($toList -and @($toList).Count -gt 0) { $mail.To = (@($toList) -join ';') }
        if ($ccList -and @($ccList).Count -gt 0) { $mail.CC = (@($ccList) -join ';') }

        # 작성한 답장 본문을 원본 인용문 위(<body> 태그 바로 뒤)에 삽입합니다.
        # <html> 앞에 붙이면 문서 구조가 깨져 Outlook 이 기본 글꼴로 렌더링합니다.
        $existing = [string]$mail.HTMLBody
        $bm = [regex]::Match($existing, '<body[^>]*>', 'IgnoreCase')
        if ($bm.Success) {
            $idx = $bm.Index + $bm.Length
            $mail.HTMLBody = $existing.Substring(0, $idx) + $styledBody + $existing.Substring($idx)
        } else {
            $mail.HTMLBody = $styledBody + $existing
        }
    } else {
        throw "알 수 없는 mode 입니다: $mode"
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
