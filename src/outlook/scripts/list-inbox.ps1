# 받은 편지함(Inbox)의 메일 목록을 가져옵니다.
# Request:  { scope: 'all' | 'unread' | 'today', maxCount: number }
# Response: { items: MailSummary[], count: number }
param(
    [Parameter(Mandatory)][string]$RequestPath,
    [Parameter(Mandatory)][string]$ResponsePath
)

. "$PSScriptRoot\_common.ps1"

try {
    $req = Read-Request $RequestPath
    $scope = Get-Prop $req 'scope' 'all'
    $maxCount = [int](Get-Prop $req 'maxCount' 200)

    $app = Get-OutlookApp
    $ns = Get-OutlookNamespace $app
    $inbox = Get-InboxFolder $ns

    $items = $inbox.Items
    $items.Sort('[ReceivedTime]', $true)

    if ($scope -eq 'unread') {
        $items = $items.Restrict('[Unread] = true')
    }

    $todayStart = (Get-Date).Date

    $results = New-Object System.Collections.Generic.List[object]
    foreach ($item in $items) {
        if ($results.Count -ge $maxCount) { break }
        if (-not (Get-ItemType ([int]$item.Class))) { continue }
        if ($scope -eq 'today' -and $item.ReceivedTime -lt $todayStart) {
            # 받은 시간 내림차순 정렬이므로 오늘보다 이전 메일이 나오면 더 볼 필요가 없습니다.
            break
        }
        $results.Add((ConvertTo-MailSummary $item $inbox)) | Out-Null
    }

    $data = [ordered]@{ items = $results; count = $results.Count }
    Write-JsonResult -ResponsePath $ResponsePath -Ok $true -Data $data
} catch {
    Write-JsonResult -ResponsePath $ResponsePath -Ok $false -ErrorMessage $_.Exception.Message
    exit 1
}
