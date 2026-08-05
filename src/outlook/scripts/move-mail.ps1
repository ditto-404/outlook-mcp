# 메일을 받은 편지함 하위의 지정한 폴더로 이동합니다. 폴더가 없으면 생성합니다.
# Request:  { entryId: string, storeId?: string, targetPath: string, createIfMissing?: bool }
# Response: { newEntryId: string, folderPath: string }
param(
    [Parameter(Mandatory)][string]$RequestPath,
    [Parameter(Mandatory)][string]$ResponsePath
)

. "$PSScriptRoot\_common.ps1"

try {
    $req = Read-Request $RequestPath
    $entryId = Get-Prop $req 'entryId'
    $storeId = Get-Prop $req 'storeId'
    $targetPath = Get-Prop $req 'targetPath'
    $createIfMissing = [bool](Get-Prop $req 'createIfMissing' $true)

    if (-not $entryId) { throw 'entryId 파라미터가 필요합니다.' }
    if (-not $targetPath) { throw 'targetPath 파라미터가 필요합니다.' }

    $app = Get-OutlookApp
    $ns = Get-OutlookNamespace $app
    $inbox = Get-InboxFolder $ns

    $item = Get-ItemByIds $ns $entryId $storeId
    $targetFolder = Resolve-FolderPath -Root $inbox -RelativePath $targetPath -CreateIfMissing $createIfMissing

    $moved = $item.Move($targetFolder)

    $data = [ordered]@{ newEntryId = $moved.EntryID; folderPath = $targetPath }
    Write-JsonResult -ResponsePath $ResponsePath -Ok $true -Data $data
} catch {
    Write-JsonResult -ResponsePath $ResponsePath -Ok $false -ErrorMessage $_.Exception.Message
    exit 1
}
