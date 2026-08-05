# 받은 편지함 하위의 폴더 트리를 조회합니다 (분류 결과 확인/디버깅용).
# Request:  {}
# Response: FolderNode  { name, path, itemCount, unreadCount, children: FolderNode[] }
param(
    [Parameter(Mandatory)][string]$RequestPath,
    [Parameter(Mandatory)][string]$ResponsePath
)

. "$PSScriptRoot\_common.ps1"

function ConvertTo-FolderNode {
    param($Folder, [string]$ParentPath)

    $path = if ($ParentPath) { "$ParentPath/$($Folder.Name)" } else { $Folder.Name }

    $unread = 0
    try { $unread = $Folder.UnReadItemCount } catch {}

    $children = New-Object System.Collections.Generic.List[object]
    foreach ($sub in $Folder.Folders) {
        $children.Add((ConvertTo-FolderNode -Folder $sub -ParentPath $path)) | Out-Null
    }

    return [ordered]@{
        name        = $Folder.Name
        path        = $path
        itemCount   = $Folder.Items.Count
        unreadCount = $unread
        children    = $children
    }
}

try {
    $app = Get-OutlookApp
    $ns = Get-OutlookNamespace $app
    $inbox = Get-InboxFolder $ns

    $data = ConvertTo-FolderNode -Folder $inbox -ParentPath ''
    Write-JsonResult -ResponsePath $ResponsePath -Ok $true -Data $data
} catch {
    Write-JsonResult -ResponsePath $ResponsePath -Ok $false -ErrorMessage $_.Exception.Message
    exit 1
}
