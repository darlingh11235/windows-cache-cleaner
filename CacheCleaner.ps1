[CmdletBinding()]
param()

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

$script:Categories = @(
    [pscustomobject]@{
        Id = 'UserTemp'
        Name = '当前用户临时文件'
        Root = [System.IO.Path]::GetTempPath()
        Mode = 'Recursive'
        Purpose = '应用程序安装、运行或更新时暂存的临时文件。'
        Impact = '通常可以重新生成；正在使用的文件会被跳过。清理前建议先关闭应用。'
    },
    [pscustomobject]@{
        Id = 'WindowsTemp'
        Name = 'Windows 临时文件'
        Root = (Join-Path $env:windir 'Temp')
        Mode = 'Recursive'
        Purpose = 'Windows 和部分安装程序使用的临时工作文件。'
        Impact = '多数可重新生成；部分文件需要管理员权限或正被系统使用，无法删除时会跳过。'
    },
    [pscustomobject]@{
        Id = 'ExplorerCache'
        Name = '缩略图与图标缓存'
        Root = (Join-Path $env:LOCALAPPDATA 'Microsoft\Windows\Explorer')
        Mode = 'ExplorerCache'
        Purpose = '资源管理器保存的图片/视频缩略图和图标预览数据库。'
        Impact = 'Windows 会重新生成缓存；清理后首次浏览图片文件夹时，预览可能短暂变慢。'
    },
    [pscustomobject]@{
        Id = 'DirectXCache'
        Name = 'DirectX 着色器缓存'
        Root = (Join-Path $env:LOCALAPPDATA 'D3DSCache')
        Mode = 'Recursive'
        Purpose = '图形应用和游戏保存的已编译着色器缓存。'
        Impact = '可由应用重新生成；清理后部分游戏首次启动或进入场景时可能短暂卡顿。'
    },
    [pscustomobject]@{
        Id = 'DeliveryOptimization'
        Name = 'Windows 更新传递缓存'
        Root = (Join-Path $env:ProgramData 'Microsoft\Windows\DeliveryOptimization\Cache')
        Mode = 'Recursive'
        Purpose = 'Windows 更新“传递优化”用于暂存、共享更新内容的缓存。'
        Impact = 'Windows 可再次下载需要的内容；清理不会卸载已安装更新，但可能增加后续下载流量。'
    }
)

$script:ScanResults = @{}
$script:Form = New-Object System.Windows.Forms.Form
$script:Form.Text = 'Windows 缓存清理工具'
$script:Form.StartPosition = 'CenterScreen'
$script:Form.Size = New-Object System.Drawing.Size(1180, 700)
$script:Form.MinimumSize = New-Object System.Drawing.Size(900, 600)
$script:Form.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9)

$layout = New-Object System.Windows.Forms.TableLayoutPanel
$layout.Dock = 'Fill'
$layout.ColumnCount = 1
$layout.RowCount = 5
$layout.Padding = New-Object System.Windows.Forms.Padding(12)
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 112))) | Out-Null
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 40))) | Out-Null
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 44))) | Out-Null
$script:Form.Controls.Add($layout)

$intro = New-Object System.Windows.Forms.Label
$intro.AutoSize = $true
$intro.MaximumSize = New-Object System.Drawing.Size(1120, 0)
$intro.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 8)
$intro.Text = '先扫描并查看每类缓存的用途与影响，再勾选想清理的类别。工具默认不会勾选任何项目；清理前还会再次列出所选内容供确认。'
$layout.Controls.Add($intro, 0, 0)

$script:List = New-Object System.Windows.Forms.ListView
$script:List.Dock = 'Fill'
$script:List.View = [System.Windows.Forms.View]::Details
$script:List.CheckBoxes = $true
$script:List.FullRowSelect = $true
$script:List.GridLines = $true
$script:List.MultiSelect = $false
$script:List.HideSelection = $false
[void]$script:List.Columns.Add('缓存类别', 180)
[void]$script:List.Columns.Add('缓存位置', 470)
[void]$script:List.Columns.Add('预计可清理', 105)
[void]$script:List.Columns.Add('文件数', 75)
[void]$script:List.Columns.Add('扫描状态', 180)
$layout.Controls.Add($script:List, 0, 1)

$script:Details = New-Object System.Windows.Forms.TextBox
$script:Details.Dock = 'Fill'
$script:Details.Multiline = $true
$script:Details.ReadOnly = $true
$script:Details.ScrollBars = 'Vertical'
$script:Details.BackColor = [System.Drawing.SystemColors]::Window
$script:Details.Margin = New-Object System.Windows.Forms.Padding(0, 8, 0, 4)
$script:Details.Text = '选择上方任一类别以查看详细说明。'
$layout.Controls.Add($script:Details, 0, 2)

$progressPanel = New-Object System.Windows.Forms.TableLayoutPanel
$progressPanel.Dock = 'Fill'
$progressPanel.ColumnCount = 2
$progressPanel.RowCount = 1
$progressPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$progressPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 180))) | Out-Null
$script:Status = New-Object System.Windows.Forms.Label
$script:Status.Dock = 'Fill'
$script:Status.TextAlign = 'MiddleLeft'
$script:Status.Text = '尚未扫描。'
$progressPanel.Controls.Add($script:Status, 0, 0)
$script:Progress = New-Object System.Windows.Forms.ProgressBar
$script:Progress.Dock = 'Fill'
$script:Progress.Style = 'Continuous'
$progressPanel.Controls.Add($script:Progress, 1, 0)
$layout.Controls.Add($progressPanel, 0, 3)

$buttons = New-Object System.Windows.Forms.FlowLayoutPanel
$buttons.Dock = 'Fill'
$buttons.FlowDirection = 'RightToLeft'
$buttons.WrapContents = $false
$buttons.Padding = New-Object System.Windows.Forms.Padding(0, 5, 0, 0)
$script:DeleteButton = New-Object System.Windows.Forms.Button
$script:DeleteButton.Text = '清理已勾选项目…'
$script:DeleteButton.AutoSize = $true
$script:DeleteButton.Enabled = $false
$script:DeleteButton.Margin = New-Object System.Windows.Forms.Padding(8, 0, 0, 0)
$script:ScanButton = New-Object System.Windows.Forms.Button
$script:ScanButton.Text = '重新扫描'
$script:ScanButton.AutoSize = $true
$script:ScanButton.Margin = New-Object System.Windows.Forms.Padding(8, 0, 0, 0)
$selectNoneButton = New-Object System.Windows.Forms.Button
$selectNoneButton.Text = '取消全选'
$selectNoneButton.AutoSize = $true
$selectNoneButton.Margin = New-Object System.Windows.Forms.Padding(8, 0, 0, 0)
$selectAllButton = New-Object System.Windows.Forms.Button
$selectAllButton.Text = '全选'
$selectAllButton.AutoSize = $true
$selectAllButton.Margin = New-Object System.Windows.Forms.Padding(8, 0, 0, 0)
[void]$buttons.Controls.Add($script:DeleteButton)
[void]$buttons.Controls.Add($script:ScanButton)
[void]$buttons.Controls.Add($selectNoneButton)
[void]$buttons.Controls.Add($selectAllButton)
$layout.Controls.Add($buttons, 0, 4)

function Format-ByteCount {
    param([long]$Bytes)
    if ($Bytes -ge 1GB) { return ('{0:N2} GB' -f ($Bytes / 1GB)) }
    if ($Bytes -ge 1MB) { return ('{0:N2} MB' -f ($Bytes / 1MB)) }
    if ($Bytes -ge 1KB) { return ('{0:N1} KB' -f ($Bytes / 1KB)) }
    return ('{0} B' -f $Bytes)
}

function Get-CacheFiles {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][string]$Mode
    )

    $files = New-Object 'System.Collections.Generic.List[object]'
    $errors = 0
    if (-not (Test-Path -LiteralPath $Root -PathType Container)) {
        return [pscustomobject]@{ Files = @(); Errors = 0; Missing = $true }
    }

    if ($Mode -eq 'ExplorerCache') {
        try {
            foreach ($item in (Get-ChildItem -LiteralPath $Root -Force -File -ErrorAction Stop)) {
                if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { continue }
                if ($item.Name -match '^(thumbcache|iconcache).*\.db$') {
                    $files.Add([pscustomobject]@{ Path = $item.FullName; Length = [long]$item.Length })
                }
            }
        }
        catch { $errors++ }
        return [pscustomobject]@{ Files = @($files.ToArray()); Errors = $errors; Missing = $false }
    }

    $pending = New-Object 'System.Collections.Generic.Stack[string]'
    $pending.Push([System.IO.Path]::GetFullPath($Root))
    while ($pending.Count -gt 0) {
        $directory = $pending.Pop()
        try { $children = @(Get-ChildItem -LiteralPath $directory -Force -ErrorAction Stop) }
        catch { $errors++; continue }
        foreach ($item in $children) {
            if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { continue }
            if ($item.PSIsContainer) {
                $pending.Push($item.FullName)
            }
            else {
                $files.Add([pscustomobject]@{ Path = $item.FullName; Length = [long]$item.Length })
            }
        }
    }
    return [pscustomobject]@{ Files = @($files.ToArray()); Errors = $errors; Missing = $false }
}

function Set-RowSubItem {
    param([System.Windows.Forms.ListViewItem]$Row, [int]$Index, [string]$Value)
    while ($Row.SubItems.Count -le $Index) { [void]$Row.SubItems.Add('') }
    $Row.SubItems[$Index].Text = $Value
}

function Update-RowScan {
    param([pscustomobject]$Category, [pscustomobject]$Result)
    $row = $script:List.Items | Where-Object { $_.Tag -eq $Category.Id } | Select-Object -First 1
    if (-not $row) { return }
    if ($Result.Missing) {
        Set-RowSubItem $row 2 '0 B'
        Set-RowSubItem $row 3 '0'
        Set-RowSubItem $row 4 '路径不存在'
        return
    }
    $total = [long]0
    foreach ($file in $Result.Files) { $total += [long]$file.Length }
    Set-RowSubItem $row 2 (Format-ByteCount $total)
    Set-RowSubItem $row 3 ([string]$Result.Files.Count)
    if ($Result.Errors -gt 0) {
        Set-RowSubItem $row 4 ('部分目录无法读取：{0}' -f $Result.Errors)
    }
    else {
        Set-RowSubItem $row 4 '扫描完成'
    }
}

function Invoke-Scan {
    $script:ScanButton.Enabled = $false
    $script:DeleteButton.Enabled = $false
    $script:Progress.Value = 0
    $script:Progress.Maximum = $script:Categories.Count
    $script:ScanResults = @{}
    $index = 0
    foreach ($category in $script:Categories) {
        $index++
        $script:Status.Text = "正在扫描：$($category.Name)"
        $row = $script:List.Items | Where-Object { $_.Tag -eq $category.Id } | Select-Object -First 1
        if ($row) {
            Set-RowSubItem $row 2 '扫描中…'
            Set-RowSubItem $row 3 '—'
            Set-RowSubItem $row 4 '扫描中…'
        }
        [System.Windows.Forms.Application]::DoEvents()
        $result = Get-CacheFiles -Root $category.Root -Mode $category.Mode
        $script:ScanResults[$category.Id] = $result
        Update-RowScan -Category $category -Result $result
        $script:Progress.Value = $index
        [System.Windows.Forms.Application]::DoEvents()
    }
    $script:Status.Text = '扫描完成。勾选要清理的类别后，点击“清理已勾选项目…”。'
    $script:ScanButton.Enabled = $true
    $script:DeleteButton.Enabled = $true
}

function Test-PathWithinRoot {
    param([string]$Root, [string]$Path)
    try {
        $rootFull = [System.IO.Path]::GetFullPath($Root).TrimEnd([char[]]@('\', '/')) + [System.IO.Path]::DirectorySeparatorChar
        $pathFull = [System.IO.Path]::GetFullPath($Path)
        return $pathFull.StartsWith($rootFull, [System.StringComparison]::OrdinalIgnoreCase)
    }
    catch { return $false }
}

function Get-CheckedCategories {
    $ids = @()
    foreach ($row in $script:List.Items) {
        if ($row.Checked) { $ids += [string]$row.Tag }
    }
    return @($script:Categories | Where-Object { $ids -contains $_.Id })
}

foreach ($category in $script:Categories) {
    $row = New-Object System.Windows.Forms.ListViewItem($category.Name)
    [void]$row.SubItems.Add($category.Root)
    [void]$row.SubItems.Add('尚未扫描')
    [void]$row.SubItems.Add('—')
    [void]$row.SubItems.Add('尚未扫描')
    $row.Tag = $category.Id
    $row.Checked = $false
    [void]$script:List.Items.Add($row)
}

$script:List.add_SelectedIndexChanged({
    if ($script:List.SelectedItems.Count -eq 0) { return }
    $id = [string]$script:List.SelectedItems[0].Tag
    $category = $script:Categories | Where-Object { $_.Id -eq $id } | Select-Object -First 1
    if ($category) {
        $script:Details.Text = "用途：$($category.Purpose)`r`n清理影响：$($category.Impact)`r`n位置：$($category.Root)"
    }
})

$script:ScanButton.add_Click({ Invoke-Scan })
$selectAllButton.add_Click({ foreach ($row in $script:List.Items) { $row.Checked = $true } })
$selectNoneButton.add_Click({ foreach ($row in $script:List.Items) { $row.Checked = $false } })

$script:DeleteButton.add_Click({
    $selected = @(Get-CheckedCategories)
    if ($selected.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show('请先勾选至少一个缓存类别。', '尚未选择项目', 'OK', 'Information')
        return
    }
    $unscanned = @($selected | Where-Object { -not $script:ScanResults.ContainsKey($_.Id) })
    if ($unscanned.Count -gt 0) {
        [void][System.Windows.Forms.MessageBox]::Show('请先完成扫描，再清理所选类别。', '需要先扫描', 'OK', 'Information')
        return
    }

    $lines = New-Object 'System.Collections.Generic.List[string]'
    $filesToDelete = New-Object 'System.Collections.Generic.List[object]'
    foreach ($category in $selected) {
        $result = $script:ScanResults[$category.Id]
        $sum = [long]0
        foreach ($file in $result.Files) {
            $sum += [long]$file.Length
            $filesToDelete.Add([pscustomobject]@{ Category = $category; File = $file })
        }
        $lines.Add(('• {0}：{1} 个文件，约 {2}' -f $category.Name, $result.Files.Count, (Format-ByteCount $sum)))
    }
    if ($filesToDelete.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show('所选类别没有找到可清理的文件。', '没有可清理内容', 'OK', 'Information')
        return
    }
    $prompt = "将尝试删除以下扫描到的缓存文件：`r`n`r`n$($lines -join "`r`n")`r`n`r`n正在使用或无权限的文件会跳过。确认继续吗？"
    $answer = [System.Windows.Forms.MessageBox]::Show($prompt, '确认清理', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning, [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
    if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $script:ScanButton.Enabled = $false
    $script:DeleteButton.Enabled = $false
    $script:Progress.Value = 0
    $script:Progress.Maximum = [Math]::Max(1, $filesToDelete.Count)
    $deleted = 0
    $skipped = 0
    $position = 0
    foreach ($entry in $filesToDelete) {
        $position++
        $path = [string]$entry.File.Path
        $root = [string]$entry.Category.Root
        $script:Status.Text = "正在清理：$($entry.Category.Name)（$position / $($filesToDelete.Count)）"
        if (-not (Test-PathWithinRoot -Root $root -Path $path)) {
            $skipped++
            continue
        }
        try {
            $item = Get-Item -LiteralPath $path -Force -ErrorAction Stop
            if ($item.PSIsContainer -or (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0)) {
                $skipped++
            }
            else {
                Remove-Item -LiteralPath $path -Force -ErrorAction Stop
                $deleted++
            }
        }
        catch { $skipped++ }
        $script:Progress.Value = [Math]::Min($position, $script:Progress.Maximum)
        if (($position % 30) -eq 0) { [System.Windows.Forms.Application]::DoEvents() }
    }
    $script:Status.Text = "清理完成：删除 $deleted 个文件，跳过 $skipped 个。正在更新扫描结果…"
    foreach ($row in $script:List.Items) { $row.Checked = $false }
    [System.Windows.Forms.Application]::DoEvents()
    Invoke-Scan
    [void][System.Windows.Forms.MessageBox]::Show("清理完成。`r`n成功删除：$deleted 个文件`r`n跳过或无法删除：$skipped 个文件`r`n`r`n扫描结果已更新。", '清理结果', 'OK', 'Information')
})

$script:Form.add_Shown({
    if ($script:List.Items.Count -gt 0) { $script:List.Items[0].Selected = $true }
})

[void]$script:Form.ShowDialog()
