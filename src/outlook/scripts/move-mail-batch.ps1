# 여러 메일을 한 번의 Outlook 연결(한 번의 PowerShell 프로세스)로 일괄 이동합니다.
# move-mail.ps1 을 건마다 반복 호출하면 매번 새 PowerShell 프로세스가 뜨느라 느리므로,
# organize_mail 처럼 대량 이동이 필요한 경우 이 스크립트를 사용합니다.
#
# Request:  { moves: [{ entryId: string, storeId?: string, targetPath: string }, ...], rootFolder?: 'inbox' | 'sent' }
# Response: { moved: [{ entryId, targetPath, newEntryId }], errors: [{ entryId, targetPath, error }] }
param(
    [Parameter(Mandatory)][string]$RequestPath,
    [Parameter(Mandatory)][string]$ResponsePath
)

. "$PSScriptRoot\_common.ps1"

try {
    $req = Read-Request $RequestPath
    $moves = @(Get-Prop $req 'moves' @())
    $rootFolderName = Get-Prop $req 'rootFolder' 'inbox'

    $app = Get-OutlookApp
    $ns = Get-OutlookNamespace $app
    $root = Get-RootFolder $ns $rootFolderName

    # 같은 targetPath 를 여러 번 Resolve 하지 않도록 캐시한다 (폴더 생성 포함).
    $folderCache = @{}

    $moved = New-Object System.Collections.Generic.List[object]
    $errors = New-Object System.Collections.Generic.List[object]

    foreach ($m in $moves) {
        $entryId = Get-Prop $m 'entryId'
        $storeId = Get-Prop $m 'storeId'
        $targetPath = Get-Prop $m 'targetPath'
        try {
            if (-not $folderCache.ContainsKey($targetPath)) {
                $folderCache[$targetPath] = Resolve-FolderPath -Root $root -RelativePath $targetPath -CreateIfMissing $true
            }
            $folder = $folderCache[$targetPath]
            $item = Get-ItemByIds $ns $entryId $storeId
            $movedItem = $item.Move($folder)
            $moved.Add([ordered]@{ entryId = $entryId; targetPath = $targetPath; newEntryId = $movedItem.EntryID }) | Out-Null
        } catch {
            $errors.Add([ordered]@{ entryId = $entryId; targetPath = $targetPath; error = $_.Exception.Message }) | Out-Null
        }
    }

    $data = [ordered]@{ moved = $moved; errors = $errors }
    Write-JsonResult -ResponsePath $ResponsePath -Ok $true -Data $data
} catch {
    Write-JsonResult -ResponsePath $ResponsePath -Ok $false -ErrorMessage $_.Exception.Message
    exit 1
}
