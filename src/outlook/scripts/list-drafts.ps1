# 임시보관함(Drafts)에 저장된 초안 목록을 조회합니다 (완성되지 않은 답장/새 메일 포함).
# 제목/키워드로 필터링할 수 있으며, 최근에 수정한 초안이 먼저 오도록 정렬합니다.
# Request:  { subject?: string, keyword?: string, maxCount?: number }
# Response: { items: DraftSummary[] }
param(
    [Parameter(Mandatory)][string]$RequestPath,
    [Parameter(Mandatory)][string]$ResponsePath
)

. "$PSScriptRoot\_common.ps1"

try {
    $req = Read-Request $RequestPath

    $subjectRaw = Get-Prop $req 'subject'
    $subjectFilter = if ($subjectRaw) { ([string]$subjectRaw).ToLower() } else { $null }

    $keywordRaw = Get-Prop $req 'keyword'
    $keywordFilter = if ($keywordRaw) { ([string]$keywordRaw).ToLower() } else { $null }

    $maxCount = [int](Get-Prop $req 'maxCount' 25)

    $app = Get-OutlookApp
    $ns = Get-OutlookNamespace $app
    $drafts = Get-DraftsFolder $ns

    $items = $drafts.Items
    try { $items.Sort('[LastModificationTime]', $true) } catch {}

    $results = New-Object System.Collections.Generic.List[object]
    foreach ($item in $items) {
        if ($results.Count -ge $maxCount) { break }
        if ([int]$item.Class -ne $script:olMail) { continue }

        $subject = ''
        try { $subject = [string]$item.Subject } catch {}
        if ($subjectFilter -and -not ($subject.ToLower().Contains($subjectFilter))) { continue }

        if ($keywordFilter) {
            $body = ''
            try { $body = [string]$item.Body } catch {}
            $to = ''
            try { $to = [string]$item.To } catch {}
            $haystack = ($subject + ' ' + $body + ' ' + $to).ToLower()
            if (-not $haystack.Contains($keywordFilter)) { continue }
        }

        $results.Add((ConvertTo-DraftSummary $item)) | Out-Null
    }

    $data = [ordered]@{ items = $results; totalCount = $drafts.Items.Count }
    Write-JsonResult -ResponsePath $ResponsePath -Ok $true -Data $data
} catch {
    Write-JsonResult -ResponsePath $ResponsePath -Ok $false -ErrorMessage $_.Exception.Message
    exit 1
}
