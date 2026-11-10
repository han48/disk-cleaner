Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic

[System.Windows.Forms.Application]::EnableVisualStyles()

$script:ScanJob = $null
$script:CleanupJob = $null
$script:CleanupResult = $null
$script:IgnoreFilePath = Join-Path $PSScriptRoot '.ignores'
$script:IgnoreRules = @()
$script:IgnoreFileContents = ''
$script:ScanIgnoreFileContents = ''
$script:ScanCandidates = New-Object 'System.Collections.Generic.List[object]'
$script:ScanRecords = @()
$script:ApplicationRoot = [System.IO.Path]::GetFullPath($PSScriptRoot).TrimEnd('\')
$script:ScanCachePath = Join-Path $script:ApplicationRoot 'DiskClean.scan-cache.json'
$script:UpdatingTreeChecks = $false

function Get-IgnoreRules {
    if (-not (Test-Path -LiteralPath $script:IgnoreFilePath -PathType Leaf)) {
        throw "Ignore file not found: $script:IgnoreFilePath"
    }

    $contents = [System.IO.File]::ReadAllText($script:IgnoreFilePath)
    $rules = New-Object 'System.Collections.Generic.List[object]'
    foreach ($rawLine in ($contents -split "\r?\n")) {
        $entry = $rawLine.Trim()
        if (-not $entry -or $entry.StartsWith('#')) { continue }

        if ([System.IO.Path]::IsPathRooted($entry)) {
            $normalizedPath = [System.IO.Path]::GetFullPath($entry).TrimEnd('\')
            if (-not $normalizedPath -or $normalizedPath -match '^[A-Za-z]:$') {
                throw "Drive roots cannot be used as ignore entries: $entry"
            }
            $rules.Add([pscustomobject]@{ Kind = 'Path'; Value = $normalizedPath })
        } elseif ($entry.Contains('\') -or $entry.Contains('/')) {
            throw "Use a folder name or an absolute path in .ignores: $entry"
        } else {
            $rules.Add([pscustomobject]@{ Kind = 'Name'; Value = $entry })
        }
    }
    if ($script:ApplicationRoot -notmatch '^[A-Za-z]:$') {
        $rules.Add([pscustomobject]@{
            Kind = 'Path'
            Value = $script:ApplicationRoot
        })
    }
    $script:IgnoreFileContents = $contents
    return @($rules.ToArray())
}

function Test-PathAtOrBelow([string]$Path, [string]$Root) {
    $normalizedRoot = $Root.TrimEnd('\')
    return $Path.Equals($normalizedRoot, [System.StringComparison]::OrdinalIgnoreCase) -or
        $Path.StartsWith($normalizedRoot + '\', [System.StringComparison]::OrdinalIgnoreCase)
}

function Test-MatchesIgnoreRule([string]$Path, $Rules) {
    $pathComponents = $Path.TrimEnd('\').Split(
        [char[]]@('\', '/'),
        [System.StringSplitOptions]::RemoveEmptyEntries
    )
    foreach ($rule in $Rules) {
        if ($rule.Kind -eq 'Name' -and $pathComponents -contains [string]$rule.Value) {
            return $true
        }
        if ($rule.Kind -eq 'Path' -and (Test-PathAtOrBelow $Path ([string]$rule.Value))) {
            return $true
        }
    }
    return $false
}

function Get-CurrentFixedDrives {
    @([System.IO.DriveInfo]::GetDrives() |
        Where-Object { $_.DriveType -eq [System.IO.DriveType]::Fixed } |
        ForEach-Object { $_.Name.TrimEnd('\').ToUpperInvariant() } |
        Sort-Object)
}

function Format-Size([int64]$Bytes) {
    if ($Bytes -ge 1TB) { return '{0:N1} TB' -f ($Bytes / 1TB) }
    if ($Bytes -ge 1GB) { return '{0:N1} GB' -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return '{0:N0} MB' -f ($Bytes / 1MB) }
    return '{0:N0} KB' -f ($Bytes / 1KB)
}

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Disk Clean'
$form.Size = New-Object System.Drawing.Size(1000, 640)
$form.MinimumSize = New-Object System.Drawing.Size(900, 420)
$form.StartPosition = 'CenterScreen'

$topPanel = New-Object System.Windows.Forms.Panel
$topPanel.Dock = 'Top'
$topPanel.Height = 116
$form.Controls.Add($topPanel)

$scanButton = New-Object System.Windows.Forms.Button
$scanButton.Text = 'Load / Scan'
$scanButton.Location = New-Object System.Drawing.Point(12, 12)
$scanButton.Size = New-Object System.Drawing.Size(110, 30)
$topPanel.Controls.Add($scanButton)

$thresholdLabel = New-Object System.Windows.Forms.Label
$thresholdLabel.Text = 'Minimum folder size (MB):'
$thresholdLabel.AutoSize = $true
$thresholdLabel.Location = New-Object System.Drawing.Point(140, 19)
$topPanel.Controls.Add($thresholdLabel)

$thresholdInput = New-Object System.Windows.Forms.NumericUpDown
$thresholdInput.Minimum = 1
$thresholdInput.Maximum = 1048576
$thresholdInput.Value = 500
$thresholdInput.Location = New-Object System.Drawing.Point(300, 15)
$thresholdInput.Size = New-Object System.Drawing.Size(100, 24)
$topPanel.Controls.Add($thresholdInput)

$deleteButton = New-Object System.Windows.Forms.Button
$deleteButton.Text = 'Move selected to Recycle Bin'
$deleteButton.Location = New-Object System.Drawing.Point(420, 12)
$deleteButton.Size = New-Object System.Drawing.Size(205, 30)
$deleteButton.Enabled = $false
$topPanel.Controls.Add($deleteButton)

$forceScanButton = New-Object System.Windows.Forms.Button
$forceScanButton.Text = 'Force scan'
$forceScanButton.Location = New-Object System.Drawing.Point(640, 12)
$forceScanButton.Size = New-Object System.Drawing.Size(105, 30)
$topPanel.Controls.Add($forceScanButton)

$sortLabel = New-Object System.Windows.Forms.Label
$sortLabel.Text = 'Sort:'
$sortLabel.AutoSize = $true
$sortLabel.Location = New-Object System.Drawing.Point(760, 19)
$topPanel.Controls.Add($sortLabel)

$sortCombo = New-Object System.Windows.Forms.ComboBox
$sortCombo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
[void]$sortCombo.Items.Add('Largest first')
[void]$sortCombo.Items.Add('Name A-Z')
$sortCombo.SelectedIndex = 0
$sortCombo.Location = New-Object System.Drawing.Point(800, 15)
$sortCombo.Size = New-Object System.Drawing.Size(150, 24)
$topPanel.Controls.Add($sortCombo)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = 'Choose a minimum size and scan fixed drives.'
$statusLabel.AutoSize = $false
$statusLabel.Size = New-Object System.Drawing.Size(850, 18)
$statusLabel.Anchor = 'Top, Left, Right'
$statusLabel.Location = New-Object System.Drawing.Point(12, 55)
$topPanel.Controls.Add($statusLabel)
$statusToolTip = New-Object System.Windows.Forms.ToolTip

$warningLabel = New-Object System.Windows.Forms.Label
$warningLabel.Text = 'Right-click a folder to ignore it; Windows and the app folder are always excluded.'
$warningLabel.ForeColor = [System.Drawing.Color]::DarkRed
$warningLabel.AutoSize = $true
$warningLabel.Location = New-Object System.Drawing.Point(12, 73)
$topPanel.Controls.Add($warningLabel)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(12, 96)
$progressBar.Size = New-Object System.Drawing.Size(850, 12)
$progressBar.Anchor = 'Top, Left, Right'
$progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Blocks
$progressBar.Visible = $false
$topPanel.Controls.Add($progressBar)

$tree = New-Object System.Windows.Forms.TreeView
$tree.Dock = 'Fill'
$tree.CheckBoxes = $true
$tree.HideSelection = $false
$tree.ShowNodeToolTips = $true
$tree.Sorted = $false
$form.Controls.Add($tree)
$topPanel.SendToBack()

$ignoreContextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$ignoreFolderMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem('Ignore this folder')
[void]$ignoreContextMenu.Items.Add($ignoreFolderMenuItem)
$ignoreCheckedFoldersMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem('Ignore checked folders')
[void]$ignoreContextMenu.Items.Add($ignoreCheckedFoldersMenuItem)

function Invoke-IgnoreFolders([string[]]$Paths) {
    try {
        $script:IgnoreRules = @(Get-IgnoreRules)
        $uniquePaths = New-Object 'System.Collections.Generic.List[string]'
        $alreadyIgnoredPaths = New-Object 'System.Collections.Generic.List[string]'
        foreach ($candidatePath in $Paths) {
            $path = [System.IO.Path]::GetFullPath($candidatePath).TrimEnd('\')
            $driveRoot = [System.IO.Path]::GetPathRoot($path).TrimEnd('\')
            if (-not $driveRoot -or $path -ieq $driveRoot) {
                throw "Drive roots cannot be ignored: $path"
            }
            if ($path -match '^[A-Za-z]:\\Windows(?:\\|$)') {
                throw "The Windows folder is always excluded: $path"
            }
            if (-not [System.IO.Directory]::Exists($path)) {
                throw "The folder no longer exists: $path"
            }
            $attributes = [System.IO.File]::GetAttributes($path)
            if (($attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
                throw "Junctions and symbolic links cannot be ignored by this app: $path"
            }

            $coveredByRule = Test-MatchesIgnoreRule $path $script:IgnoreRules
            $coveredBySelection = $false
            foreach ($existingPath in $uniquePaths) {
                if (Test-PathAtOrBelow $path $existingPath) {
                    $coveredBySelection = $true
                    break
                }
            }
            if ($coveredByRule) {
                $alreadyIgnoredPaths.Add($path)
            } elseif (-not $coveredBySelection) {
                $uniquePaths.Add($path)
            }
        }

        if ($uniquePaths.Count -eq 0) {
            Mask-ScanRecords $alreadyIgnoredPaths.ToArray() -Excluded
            try {
                Write-ScanCache
            } catch {
                $statusLabel.Text = "Ignored folders were hidden, but the saved scan could not be updated: $($_.Exception.Message)"
                [System.Windows.Forms.MessageBox]::Show(
                    $statusLabel.Text,
                    'Saved scan could not be updated',
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                ) | Out-Null
                return
            }
            [System.Windows.Forms.MessageBox]::Show(
                'Already-ignored folders have been hidden from the current results.',
                'Folders already ignored',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
            $statusLabel.Text = 'Already-ignored folders are hidden from the current results and saved scan.'
            return
        }

        $preview = ($uniquePaths | Select-Object -First 8) -join "`r`n"
        if ($uniquePaths.Count -gt 8) { $preview += "`r`n... and $($uniquePaths.Count - 8) more" }
        $answer = [System.Windows.Forms.MessageBox]::Show(
            "Ignore $($uniquePaths.Count) folder(s) and all of their contents in future scans?`r`n`r`n$preview`r`n`r`nThey will be added to .ignores, removed from the saved scan, and hidden from the current results.",
            'Ignore folders',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        $scanRulesUnchanged = $script:ScanIgnoreFileContents -ceq $script:IgnoreFileContents
        $contents = $script:IgnoreFileContents
        $prefix = ''
        if ($contents.Length -gt 0 -and $contents -notmatch '[\r\n]$') {
            $prefix = [Environment]::NewLine
        }
        $encoding = New-Object System.Text.UTF8Encoding($false)
        [System.IO.File]::AppendAllText(
            $script:IgnoreFilePath,
            $prefix + ($uniquePaths -join [Environment]::NewLine) + [Environment]::NewLine,
            $encoding
        )
        $script:IgnoreRules = @(Get-IgnoreRules)
        $pathsToMask = @($uniquePaths.ToArray()) + @($alreadyIgnoredPaths.ToArray())
        Mask-ScanRecords $pathsToMask -Excluded
        if ($scanRulesUnchanged) {
            $script:ScanIgnoreFileContents = $script:IgnoreFileContents
        }
        try {
            Write-ScanCache
        } catch {
            $statusLabel.Text = "Ignored folders are hidden from current results, but the saved scan could not be updated: $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show(
                "$($statusLabel.Text)`r`nThe next app launch will scan again if the cache no longer matches .ignores.",
                'Saved scan could not be updated',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null
            return
        }
        $statusLabel.Text = "Ignored and hid $($pathsToMask.Count) folder(s) from the current results and saved scan."
    } catch {
        $statusLabel.Text = "Unable to ignore folder(s): $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show(
            $statusLabel.Text,
            'Unable to ignore folders',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
}

$tree.Add_NodeMouseClick({
    if ($args[1].Button -eq [System.Windows.Forms.MouseButtons]::Right) {
        $tree.SelectedNode = $args[1].Node
        $ignoreContextMenu.Show($tree, $args[1].Location)
    }
})

$ignoreFolderMenuItem.Add_Click({
    $node = $tree.SelectedNode
    if ($null -eq $node) { return }
    Invoke-IgnoreFolders @([string]$node.Tag.Path)
})

$ignoreCheckedFoldersMenuItem.Add_Click({
    $checkedPaths = New-Object 'System.Collections.Generic.List[string]'
    $collectCheckedPaths = {
        param($Nodes)
        foreach ($node in $Nodes) {
            if ($node.Checked -and [bool]$node.Tag.CanDelete) {
                $hasCheckedAncestor = $false
                $parent = $node.Parent
                while ($null -ne $parent) {
                    if ($parent.Checked -and [bool]$parent.Tag.CanDelete) {
                        $hasCheckedAncestor = $true
                        break
                    }
                    $parent = $parent.Parent
                }
                if (-not $hasCheckedAncestor) {
                    $checkedPaths.Add([string]$node.Tag.Path)
                }
            }
            & $collectCheckedPaths $node.Nodes
        }
    }
    & $collectCheckedPaths $tree.Nodes
    if ($checkedPaths.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            'Check at least one safe folder in the results tree.',
            'No folders selected',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        return
    }
    Invoke-IgnoreFolders $checkedPaths.ToArray()
})

function Sort-TreeNodes($Nodes) {
    if ($sortCombo.SelectedIndex -eq 1) {
        $sortedNodes = @($Nodes | Sort-Object { $_.Text })
    } else {
        $sortedNodes = @($Nodes | Sort-Object @{ Expression = { [int64]$_.Tag.Bytes }; Descending = $true }, { $_.Text })
    }
    $Nodes.Clear()
    foreach ($node in $sortedNodes) {
        [void]$Nodes.Add($node)
        Sort-TreeNodes $node.Nodes
    }
}

function Set-TreeNodeChecks($ParentNode, [bool]$Checked) {
    foreach ($childNode in $ParentNode.Nodes) {
        if (-not $Checked -or [bool]$childNode.Tag.CanDelete) {
            $childNode.Checked = $Checked
        }
        Set-TreeNodeChecks $childNode $Checked
    }
}

function Render-ScanTree {
    $tree.BeginUpdate()
    try {
        $tree.Nodes.Clear()
        $minimumBytes = [int64]$thresholdInput.Value * 1MB
        $recordsByPath = @{}
        foreach ($record in $script:ScanRecords) {
            $recordsByPath[[string]$record.Path] = $record
        }

        $visiblePaths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        $qualifyingRecords = @($script:ScanRecords | Where-Object { [int64]$_.Bytes -ge $minimumBytes })
        foreach ($record in $qualifyingRecords) {
            $currentPath = [string]$record.Path
            while ($currentPath -match '^[A-Za-z]:\\') {
                if ($currentPath -match '^[A-Za-z]:\\$') {
                    [void]$visiblePaths.Add($currentPath)
                    break
                }
                [void]$visiblePaths.Add($currentPath.TrimEnd('\'))
                $parent = [System.IO.Directory]::GetParent($currentPath)
                if ($null -eq $parent -or $parent.FullName -eq $currentPath) { break }
                $currentPath = $parent.FullName
            }
        }

        $nodesByPath = @{}
        $orderedPaths = @($visiblePaths | Sort-Object { $_.Length }, { $_ })
        foreach ($path in $orderedPaths) {
            $record = $recordsByPath[$path]
            if ($null -ne $record) {
                $bytes = [int64]$record.Bytes
                $canDelete = [bool]$record.CanDelete -and $bytes -ge $minimumBytes
                $label = '{0}  ({1})' -f [System.IO.Path]::GetFileName($path), (Format-Size $bytes)
                if (-not [bool]$record.CanDelete) {
                    $label += '  [contains excluded folders]'
                } elseif ($bytes -lt $minimumBytes) {
                    $label += '  [expand for large subfolders]'
                }
                $tag = [pscustomobject]@{ Path = $path; Bytes = $bytes; CanDelete = $canDelete }
            } else {
                $label = if ($path -match '^[A-Za-z]:\\$') { $path } else { [System.IO.Path]::GetFileName($path) }
                $tag = [pscustomobject]@{ Path = $path; Bytes = [int64]0; CanDelete = $false }
            }

            $node = New-Object System.Windows.Forms.TreeNode($label)
            $node.Tag = $tag
            $node.ToolTipText = $path
            $nodesByPath[$path] = $node
            $parentPath = [System.IO.Directory]::GetParent($path)
            if ($null -ne $parentPath -and $nodesByPath.ContainsKey($parentPath.FullName)) {
                [void]$nodesByPath[$parentPath.FullName].Nodes.Add($node)
            } else {
                [void]$tree.Nodes.Add($node)
            }
        }

        Sort-TreeNodes $tree.Nodes
        $tree.ExpandAll()
        $deleteButton.Enabled = @($qualifyingRecords | Where-Object { [bool]$_.CanDelete }).Count -gt 0
        $statusLabel.Text = "Showing $($qualifyingRecords.Count) folder(s) at or above $([int]$thresholdInput.Value) MB from saved scan."
    } finally {
        $tree.EndUpdate()
    }
}

function Mask-ScanRecords([string[]]$Paths, [switch]$Excluded) {
    $maskRoots = New-Object 'System.Collections.Generic.List[string]'
    foreach ($path in $Paths) {
        $normalizedPath = [System.IO.Path]::GetFullPath($path).TrimEnd('\')
        $coveredByExistingRoot = $false
        foreach ($root in $maskRoots) {
            if (Test-PathAtOrBelow $normalizedPath $root) {
                $coveredByExistingRoot = $true
                break
            }
        }
        if (-not $coveredByExistingRoot) { $maskRoots.Add($normalizedPath) }
    }

    $bytesByRoot = @{}
    foreach ($root in $maskRoots) {
        $rootRecord = $script:ScanRecords |
            Where-Object { ([string]$_.Path).Equals($root, [System.StringComparison]::OrdinalIgnoreCase) } |
            Select-Object -First 1
        if ($null -ne $rootRecord) {
            $bytesByRoot[$root] = [int64]$rootRecord.Bytes
        } else {
            $bytesByRoot[$root] = [int64](($script:ScanRecords |
                Where-Object {
                    $recordPath = [string]$_.Path
                    (Test-PathAtOrBelow $recordPath $root) -and
                    [System.IO.Path]::GetDirectoryName($recordPath).TrimEnd('\').Equals(
                        $root,
                        [System.StringComparison]::OrdinalIgnoreCase
                    )
                } |
                Measure-Object -Property Bytes -Sum).Sum)
        }
    }

    $updatedRecords = New-Object 'System.Collections.Generic.List[object]'
    foreach ($record in $script:ScanRecords) {
        $recordPath = [string]$record.Path
        $isMasked = $false
        $bytesToSubtract = [int64]0
        $isAncestorOfExcludedPath = $false
        foreach ($root in $maskRoots) {
            if (Test-PathAtOrBelow $recordPath $root) {
                $isMasked = $true
                break
            }
            if (Test-PathAtOrBelow $root $recordPath) {
                $bytesToSubtract += [int64]$bytesByRoot[$root]
                if ($Excluded) { $isAncestorOfExcludedPath = $true }
            }
        }
        if ($isMasked) { continue }

        $bytes = [int64]$record.Bytes - $bytesToSubtract
        if ($bytes -lt 0) { $bytes = 0 }
        $updatedRecords.Add([pscustomobject]@{
            Path = $recordPath
            Bytes = $bytes
            CanDelete = [bool]$record.CanDelete -and -not $isAncestorOfExcludedPath
        })
    }
    $script:ScanRecords = @($updatedRecords.ToArray())
    Render-ScanTree
}

function Hide-IgnoredScanRecords {
    $ignoredPaths = @($script:ScanRecords |
        Where-Object { Test-MatchesIgnoreRule ([string]$_.Path) $script:IgnoreRules } |
        ForEach-Object { [string]$_.Path })
    if ($ignoredPaths.Count -eq 0) { return $false }
    Mask-ScanRecords $ignoredPaths -Excluded
    return $true
}

function Write-ScanCache {
    $cache = [pscustomobject]@{
        Version = 2
        ScannedAt = [DateTime]::UtcNow.ToString('o')
        ApplicationRoot = $script:ApplicationRoot
        IgnoreFileContents = $script:ScanIgnoreFileContents
        Drives = @(Get-CurrentFixedDrives)
        Folders = @($script:ScanRecords)
    }
    $temporaryPath = $script:ScanCachePath + '.tmp'
    $json = ConvertTo-Json -InputObject $cache -Depth 5
    [System.IO.File]::WriteAllText($temporaryPath, $json, (New-Object System.Text.UTF8Encoding($false)))
    Move-Item -LiteralPath $temporaryPath -Destination $script:ScanCachePath -Force
}

function Import-ScanCache {
    if (-not (Test-Path -LiteralPath $script:ScanCachePath -PathType Leaf)) { return $false }
    $cache = Get-Content -LiteralPath $script:ScanCachePath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    if ($cache.Version -ne 2) { return $false }
    if ([string]$cache.ApplicationRoot -ine $script:ApplicationRoot) { return $false }
    if ([string]$cache.IgnoreFileContents -cne $script:IgnoreFileContents) { return $false }
    $currentDrives = @(Get-CurrentFixedDrives)
    $cachedDrives = @($cache.Drives | ForEach-Object { ([string]$_).ToUpperInvariant() } | Sort-Object)
    if (($currentDrives -join '|') -cne ($cachedDrives -join '|')) { return $false }
    foreach ($record in @($cache.Folders)) {
        if (-not $record.Path -or $null -eq $record.Bytes -or $null -eq $record.CanDelete) {
            throw 'The scan cache has an invalid folder record.'
        }
    }
    $script:ScanRecords = @($cache.Folders)
    $script:ScanIgnoreFileContents = [string]$cache.IgnoreFileContents
    if (Hide-IgnoredScanRecords) {
        try {
            Write-ScanCache
        } catch {
            $statusLabel.Text = "Ignored folders were hidden, but the saved scan could not be updated: $($_.Exception.Message)"
            return $true
        }
    } else {
        Render-ScanTree
    }
    $scanTime = [DateTime]::Parse([string]$cache.ScannedAt).ToLocalTime()
    if (-not $statusLabel.Text.StartsWith('Ignored folders were hidden')) {
        $statusLabel.Text = "Loaded saved scan from $($scanTime.ToString('g')). Adjust the size threshold to reorganize results."
    }
    return $true
}

$tree.add_AfterCheck({
    if ($script:UpdatingTreeChecks) { return }
    $script:UpdatingTreeChecks = $true
    try {
        $node = $args[1].Node
        if ($node.Checked -and -not [bool]$node.Tag.CanDelete) {
            $node.Checked = $false
            $statusLabel.Text = 'This folder contains excluded data or is only a parent path; select a safe subfolder instead.'
            return
        }
        Set-TreeNodeChecks $node $node.Checked
        if (-not $node.Checked) {
            $parentNode = $node.Parent
            while ($null -ne $parentNode) {
                $parentNode.Checked = $false
                $parentNode = $parentNode.Parent
            }
        }
    } finally {
        $script:UpdatingTreeChecks = $false
    }
})

$thresholdInput.Add_ValueChanged({
    if ($script:ScanRecords.Count -gt 0) { Render-ScanTree }
})
$sortCombo.Add_SelectedIndexChanged({
    if ($script:ScanRecords.Count -gt 0) { Render-ScanTree }
})

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 500
$timer.Add_Tick({
    if ($null -ne $script:CleanupJob) {
        $cleanupMessages = @(Receive-Job -Job $script:CleanupJob -ErrorAction SilentlyContinue)
        $cleanupResult = $null
        foreach ($message in $cleanupMessages) {
            if ($message.RecordType -eq 'Progress') {
                $statusLabel.Text = 'Moving to Recycle Bin ({0}/{1}): {2}' -f `
                    $message.Index, $message.Total, $message.Path
                $statusToolTip.SetToolTip($statusLabel, [string]$message.Path)
            } elseif ($message.RecordType -eq 'Result') {
                $script:CleanupResult = $message
            }
        }

        if ($script:CleanupJob.State -in @('Completed', 'Failed', 'Stopped')) {
            $jobState = $script:CleanupJob.State
            $jobErrors = @($script:CleanupJob.ChildJobs[0].Error)
            $cleanupResult = $script:CleanupResult
            $script:CleanupResult = $null
            Remove-Job -Job $script:CleanupJob -Force
            $script:CleanupJob = $null
            $timer.Stop()
            $progressBar.Visible = $false
            $deleteButton.Text = 'Move selected to Recycle Bin'
            $deleteButton.Enabled = $true
            $scanButton.Enabled = $true
            $forceScanButton.Enabled = $true
            $thresholdInput.Enabled = $true
            $sortCombo.Enabled = $true
            $tree.Enabled = $true

            if ($jobState -ne 'Completed' -or $null -eq $cleanupResult) {
                $errorText = ($jobErrors | ForEach-Object { $_.ToString() }) -join ' '
                if (-not $errorText) { $errorText = 'The cleanup job ended without a result.' }
                $statusLabel.Text = "Cleanup failed: $errorText"
                [System.Windows.Forms.MessageBox]::Show(
                    $statusLabel.Text,
                    'Cleanup failed',
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                ) | Out-Null
                return
            }

            $deleted = [int]$cleanupResult.Deleted
            $failures = @($cleanupResult.Failures)
            $deletedPaths = @($cleanupResult.DeletedPaths)
            if ($deletedPaths.Count -gt 0) {
                Mask-ScanRecords $deletedPaths
                try {
                    Write-ScanCache
                } catch {
                    $cacheError = $_.Exception.Message
                    try {
                        if (Test-Path -LiteralPath $script:ScanCachePath -PathType Leaf) {
                            Remove-Item -LiteralPath $script:ScanCachePath -Force -ErrorAction Stop
                        }
                    } catch {
                        $cacheError += " Could not invalidate the stale cache: $($_.Exception.Message)"
                    }
                    [System.Windows.Forms.MessageBox]::Show(
                        "Cleanup completed, but the saved scan could not be updated: $cacheError`r`nUse Force scan before relying on saved results.",
                        'Saved scan needs refresh',
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Warning
                    ) | Out-Null
                }
            }
            if ($failures.Count -gt 0) {
                $failureText = ($failures | Select-Object -First 5) -join "`r`n"
                [System.Windows.Forms.MessageBox]::Show(
                    "Moved $deleted folder(s) to the Recycle Bin. Failed:`r`n$failureText",
                    'Cleanup partially completed',
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                ) | Out-Null
            } else {
                [System.Windows.Forms.MessageBox]::Show(
                    "Moved $deleted folder(s) to the Recycle Bin.",
                    'Cleanup complete',
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                ) | Out-Null
            }
            if ($deleted -gt 0) {
                $statusLabel.Text = 'Cleanup complete. Moved folders are hidden from the current results and saved scan.'
            } else {
                $statusLabel.Text = 'Cleanup finished. No folders were moved; current results are unchanged.'
            }
        }
        return
    }

    if ($null -eq $script:ScanJob) { return }
    $scanMessages = @(Receive-Job -Job $script:ScanJob -ErrorAction SilentlyContinue)
    foreach ($message in $scanMessages) {
        if ($message.RecordType -eq 'Progress') {
            $statusLabel.Text = 'Scanning {0}: {1:N0} folders, {2:N0} files checked. Current: {3}' -f `
                $message.Drive, $message.Folders, $message.Files, $message.Path
            $statusToolTip.SetToolTip($statusLabel, [string]$message.Path)
        } elseif ($message.RecordType -eq 'Folder') {
            $script:ScanCandidates.Add($message)
        }
    }

    if ($script:ScanJob.State -in @('Completed', 'Failed', 'Stopped')) {
        $jobState = $script:ScanJob.State
        $jobErrors = @($script:ScanJob.ChildJobs[0].Error)
        Remove-Job -Job $script:ScanJob -Force
        $script:ScanJob = $null
        $timer.Stop()
        $progressBar.Visible = $false

        if ($jobState -ne 'Completed') {
            $statusLabel.Text = 'Scan failed. ' + (($jobErrors | ForEach-Object { $_.ToString() }) -join ' ')
            $scanButton.Enabled = $true
            $forceScanButton.Enabled = $true
            $thresholdInput.Enabled = $true
            $sortCombo.Enabled = $true
            return
        }

        $script:ScanRecords = @($script:ScanCandidates.ToArray())
        $script:ScanCandidates.Clear()
        $script:ScanIgnoreFileContents = $script:IgnoreFileContents
        $cacheSaved = $true
        $cacheError = ''
        try {
            Write-ScanCache
        } catch {
            $cacheSaved = $false
            $cacheError = $_.Exception.Message
        }
        Render-ScanTree
        if ($cacheSaved) {
            $statusLabel.Text = "Scan complete. Saved $($script:ScanRecords.Count) folder record(s) for future use."
        } else {
            $statusLabel.Text = "Scan complete, but unable to save cache: $cacheError"
        }
        $scanButton.Enabled = $true
        $forceScanButton.Enabled = $true
        $thresholdInput.Enabled = $true
        $sortCombo.Enabled = $true
    }
})

function Start-FullScan {
    if ($null -ne $script:ScanJob) { return }
    $statusLabel.Text = 'Scanning fixed drives. This may take a while...'
    $scanButton.Enabled = $false
    $forceScanButton.Enabled = $false
    $deleteButton.Enabled = $false
    $thresholdInput.Enabled = $false
    $sortCombo.Enabled = $false
    $tree.Nodes.Clear()
    $script:ScanRecords = @()
    $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
    $progressBar.MarqueeAnimationSpeed = 25
    $progressBar.Visible = $true

    try {
        $script:IgnoreRules = @(Get-IgnoreRules)
    } catch {
        $statusLabel.Text = "Cannot read .ignores: $($_.Exception.Message)"
        $scanButton.Enabled = $true
        $forceScanButton.Enabled = $true
        $thresholdInput.Enabled = $true
        $sortCombo.Enabled = $true
        $progressBar.Visible = $false
        return
    }
    $script:ScanIgnoreFileContents = $script:IgnoreFileContents
    $ignoreRulesJson = ConvertTo-Json -InputObject @($script:IgnoreRules) -Compress
    $script:ScanCandidates.Clear()
    $script:ScanJob = Start-Job -ArgumentList $ignoreRulesJson -ScriptBlock {
        param([string]$IgnoreRulesJson)

        $ignoreRules = @(ConvertFrom-Json -InputObject $IgnoreRulesJson)

        function Test-PathAtOrBelow([string]$Path, [string]$Root) {
            $normalizedRoot = $Root.TrimEnd('\')
            return $Path.Equals($normalizedRoot, [System.StringComparison]::OrdinalIgnoreCase) -or
                $Path.StartsWith($normalizedRoot + '\', [System.StringComparison]::OrdinalIgnoreCase)
        }

        function Test-MatchesIgnoreRule([string]$Path, $Rules) {
            $pathComponents = $Path.TrimEnd('\').Split(
                [char[]]@('\', '/'),
                [System.StringSplitOptions]::RemoveEmptyEntries
            )
            foreach ($rule in $Rules) {
                if ($rule.Kind -eq 'Name' -and $pathComponents -contains [string]$rule.Value) {
                    return $true
                }
                if ($rule.Kind -eq 'Path' -and (Test-PathAtOrBelow $Path ([string]$rule.Value))) {
                    return $true
                }
            }
            return $false
        }

        $profileRoots = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        if ($env:USERPROFILE) {
            [void]$profileRoots.Add([System.IO.Path]::GetFullPath($env:USERPROFILE).TrimEnd('\'))
        }
        $usersRoot = Join-Path $env:SystemDrive 'Users'
        if ([System.IO.Directory]::Exists($usersRoot)) {
            try {
                foreach ($profile in [System.IO.Directory]::EnumerateDirectories($usersRoot)) {
                    [void]$profileRoots.Add([System.IO.Path]::GetFullPath($profile).TrimEnd('\'))
                }
            } catch { }
        }
        $safeAppDataRelativePaths = @(
            'AppData\Local\Temp',
            'AppData\Local\CrashDumps',
            'AppData\Local\D3DSCache',
            'AppData\Local\Microsoft\Windows\INetCache',
            'AppData\Local\Microsoft\Windows\WER',
            'AppData\Local\NVIDIA\DXCache',
            'AppData\Local\NVIDIA\GLCache',
            'AppData\Local\AMD\DxCache',
            'AppData\Local\AMD\GLCache'
        )
        $safeAppDataRoots = @(
            foreach ($profileRoot in $profileRoots) {
                foreach ($relativePath in $safeAppDataRelativePaths) {
                    Join-Path $profileRoot $relativePath
                }
            }
        )

        function Test-IsAppDataPath([string]$Path) {
            foreach ($profileRoot in $profileRoots) {
                $appDataRoot = Join-Path $profileRoot 'AppData'
                if (Test-PathAtOrBelow $Path $appDataRoot) { return $true }
            }
            return $false
        }

        function Test-IsSafeAppDataPath([string]$Path) {
            foreach ($root in $safeAppDataRoots) {
                if (Test-PathAtOrBelow $Path $root) { return $true }
            }
            return $false
        }

        function Test-ExcludedPath([string]$Path) {
            if ($Path -match '^[A-Za-z]:\\Windows(?:\\|$)') { return $true }
            if (Test-MatchesIgnoreRule $Path $ignoreRules) { return $true }
            if (Test-IsAppDataPath $Path) {
                foreach ($root in $safeAppDataRoots) {
                    if ((Test-PathAtOrBelow $Path $root) -or
                        (Test-PathAtOrBelow $root $Path)) {
                        return $false
                    }
                }
                return $true
            }
            return $false
        }

        function Add-UnsafeAncestors([string]$Path) {
            $ancestor = New-Object System.IO.DirectoryInfo($Path)
            while ($null -ne $ancestor -and $ancestor.FullName -match '^[A-Za-z]:\\') {
                [void]$unsafeCandidatePaths.Add($ancestor.FullName.TrimEnd('\'))
                $ancestor = $ancestor.Parent
            }
        }

        $sizeByPath = @{}
        $allDirectories = New-Object 'System.Collections.Generic.List[string]'
        $unsafeCandidatePaths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        $foldersScanned = [int64]0
        $filesScanned = [int64]0
        $lastProgressAt = [DateTime]::UtcNow
        $drives = @([System.IO.DriveInfo]::GetDrives() |
            Where-Object { $_.DriveType -eq [System.IO.DriveType]::Fixed })

        foreach ($drive in $drives) {
            $root = $drive.RootDirectory.FullName
            [pscustomobject]@{
                RecordType = 'Progress'
                Drive = $drive.Name
                Folders = $foldersScanned
                Files = $filesScanned
                Path = $root
            }
            $stack = New-Object 'System.Collections.Generic.Stack[object]'
            try {
                foreach ($entry in [System.IO.Directory]::EnumerateDirectories($root)) {
                    $stack.Push(@($entry, $false))
                }
            } catch {
                Add-UnsafeAncestors $root
                continue
            }

            while ($stack.Count -gt 0) {
                $item = $stack.Pop()
                $path = [string]$item[0]
                $expanded = [bool]$item[1]
                if (-not $expanded) {
                    $foldersScanned++
                    $now = [DateTime]::UtcNow
                    if (($foldersScanned % 100) -eq 0 -or ($now - $lastProgressAt).TotalSeconds -ge 1) {
                        [pscustomobject]@{
                            RecordType = 'Progress'
                            Drive = $drive.Name
                            Folders = $foldersScanned
                            Files = $filesScanned
                            Path = $path
                        }
                        $lastProgressAt = $now
                    }
                }

                if ($expanded) {
                    $total = [int64]0
                    try {
                        foreach ($child in [System.IO.Directory]::EnumerateDirectories($path)) {
                            if ($sizeByPath.ContainsKey($child)) { $total += [int64]$sizeByPath[$child] }
                        }
                    } catch { Add-UnsafeAncestors $path }
                    $sizeByPath[$path] = [int64]$sizeByPath[$path] + $total
                    if (-not $profileRoots.Contains($path.TrimEnd('\')) -and
                        (-not (Test-IsAppDataPath $path) -or (Test-IsSafeAppDataPath $path))) {
                        $allDirectories.Add($path)
                    }
                    continue
                }

                if (Test-ExcludedPath $path) {
                    Add-UnsafeAncestors $path
                    continue
                }
                try {
                    $attributes = [System.IO.File]::GetAttributes($path)
                    if (($attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
                        Add-UnsafeAncestors $path
                        continue
                    }
                } catch {
                    Add-UnsafeAncestors $path
                    continue
                }

                $fileBytes = [int64]0
                try {
                    foreach ($file in [System.IO.Directory]::EnumerateFiles($path)) {
                        $filesScanned++
                        $now = [DateTime]::UtcNow
                        if (($filesScanned % 1000) -eq 0 -or ($now - $lastProgressAt).TotalSeconds -ge 1) {
                            [pscustomobject]@{
                                RecordType = 'Progress'
                                Drive = $drive.Name
                                Folders = $foldersScanned
                                Files = $filesScanned
                                Path = $path
                            }
                            $lastProgressAt = $now
                        }
                        try {
                            $fileBytes += [int64]([System.IO.FileInfo]::new($file).Length)
                        } catch {
                            Add-UnsafeAncestors $path
                        }
                    }
                } catch { Add-UnsafeAncestors $path }

                $sizeByPath[$path] = $fileBytes
                $stack.Push(@($path, $true))
                try {
                    foreach ($child in [System.IO.Directory]::EnumerateDirectories($path)) {
                        $stack.Push(@($child, $false))
                    }
                } catch { Add-UnsafeAncestors $path }
            }
        }

        foreach ($path in $allDirectories) {
            [pscustomobject]@{
                RecordType = 'Folder'
                Path = $path
                Bytes = [int64]$sizeByPath[$path]
                CanDelete = -not $unsafeCandidatePaths.Contains($path)
            }
        }
    }
    $timer.Start()
}

$scanButton.Add_Click({
    if ($null -ne $script:ScanJob) { return }
    try {
        $script:IgnoreRules = @(Get-IgnoreRules)
        if (Import-ScanCache) { return }
    } catch {
        $statusLabel.Text = "Cannot load saved scan: $($_.Exception.Message)"
    }
    Start-FullScan
})

$forceScanButton.Add_Click({
    if ($null -eq $script:ScanJob) { Start-FullScan }
})

$deleteButton.Add_Click({
    $checkedNodes = New-Object 'System.Collections.Generic.List[System.Windows.Forms.TreeNode]'
    $collectCheckedNodes = {
        param($Nodes)
        foreach ($node in $Nodes) {
            if ($node.Checked -and [bool]$node.Tag.CanDelete) {
                $checkedNodes.Add($node)
            }
            & $collectCheckedNodes $node.Nodes
        }
    }
    & $collectCheckedNodes $tree.Nodes
    $checkedPaths = New-Object 'System.Collections.Generic.List[string]'
    foreach ($node in $checkedNodes) {
        $hasCheckedAncestor = $false
        $parent = $node.Parent
        while ($null -ne $parent) {
            if ($parent.Checked -and [bool]$parent.Tag.CanDelete) {
                $hasCheckedAncestor = $true
                break
            }
            $parent = $parent.Parent
        }
        if (-not $hasCheckedAncestor) {
            $checkedPaths.Add([string]$node.Tag.Path)
        }
    }
    if ($checkedPaths.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            'Check at least one safe folder in the tree.',
            'Nothing selected',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        return
    }

    $paths = @($checkedPaths.ToArray())
    $preview = ($paths | Select-Object -First 8) -join "`r`n"
    if ($paths.Count -gt 8) { $preview += "`r`n... and $($paths.Count - 8) more" }
    $answer = [System.Windows.Forms.MessageBox]::Show(
        "Move $($paths.Count) selected folder(s) and their contents to the Recycle Bin?`r`n`r`n$preview`r`n`r`nReview the paths carefully. Some application data may be in use.",
        'Confirm folder cleanup',
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    try {
        $currentIgnoreRules = @(Get-IgnoreRules)
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Cannot read .ignores: $($_.Exception.Message)",
            'Cleanup cancelled',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        return
    }
    if ($script:IgnoreFileContents -cne $script:ScanIgnoreFileContents) {
        [System.Windows.Forms.MessageBox]::Show(
            'The .ignores file changed after the scan. Scan again before deleting.',
            'Cleanup cancelled',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        $scanButton.PerformClick()
        return
    }
    $script:IgnoreRules = $currentIgnoreRules

    $pathsJson = ConvertTo-Json -InputObject @($paths) -Compress
    $ignoreRulesJson = ConvertTo-Json -InputObject @($script:IgnoreRules) -Compress
    $script:CleanupResult = $null
    $cleanupScript = {
        param([string]$PathsJson, [string]$IgnoreRulesJson)

        $paths = @(ConvertFrom-Json -InputObject $PathsJson)
        $ignoreRules = @(ConvertFrom-Json -InputObject $IgnoreRulesJson)

        function Test-PathAtOrBelow([string]$Path, [string]$Root) {
            $normalizedRoot = $Root.TrimEnd('\')
            return $Path.Equals($normalizedRoot, [System.StringComparison]::OrdinalIgnoreCase) -or
                $Path.StartsWith($normalizedRoot + '\', [System.StringComparison]::OrdinalIgnoreCase)
        }

        function Test-MatchesIgnoreRule([string]$Path, $Rules) {
            $pathComponents = $Path.TrimEnd('\').Split(
                [char[]]@('\', '/'),
                [System.StringSplitOptions]::RemoveEmptyEntries
            )
            foreach ($rule in $Rules) {
                if ($rule.Kind -eq 'Name' -and $pathComponents -contains [string]$rule.Value) {
                    return $true
                }
                if ($rule.Kind -eq 'Path' -and (Test-PathAtOrBelow $Path ([string]$rule.Value))) {
                    return $true
                }
            }
            return $false
        }

        $safeAppDataRelativePaths = @(
            'AppData\Local\Temp',
            'AppData\Local\CrashDumps',
            'AppData\Local\D3DSCache',
            'AppData\Local\Microsoft\Windows\INetCache',
            'AppData\Local\Microsoft\Windows\WER',
            'AppData\Local\NVIDIA\DXCache',
            'AppData\Local\NVIDIA\GLCache',
            'AppData\Local\AMD\DxCache',
            'AppData\Local\AMD\GLCache'
        )
        $profileRoots = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        if ($env:USERPROFILE) {
            [void]$profileRoots.Add([System.IO.Path]::GetFullPath($env:USERPROFILE).TrimEnd('\'))
        }
        $usersRoot = Join-Path $env:SystemDrive 'Users'
        if ([System.IO.Directory]::Exists($usersRoot)) {
            try {
                foreach ($profile in [System.IO.Directory]::EnumerateDirectories($usersRoot)) {
                    [void]$profileRoots.Add([System.IO.Path]::GetFullPath($profile).TrimEnd('\'))
                }
            } catch { }
        }

        $deleted = 0
        $deletedPaths = New-Object 'System.Collections.Generic.List[string]'
        $failures = New-Object 'System.Collections.Generic.List[string]'
        for ($index = 0; $index -lt $paths.Count; $index++) {
            $path = [string]$paths[$index]
            [pscustomobject]@{
                RecordType = 'Progress'
                Index = $index + 1
                Total = $paths.Count
                Path = $path
            }
            try {
                $fullPath = [System.IO.Path]::GetFullPath($path).TrimEnd('\')
                $driveRoot = [System.IO.Path]::GetPathRoot($fullPath).TrimEnd('\')
                if (-not $fullPath.StartsWith($driveRoot + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
                    throw 'The folder is not below a drive root.'
                }
                if ($fullPath -match '^[A-Za-z]:\\Windows(?:\\|$)') {
                    throw 'The Windows folder is always protected.'
                }
                if (Test-MatchesIgnoreRule $fullPath $ignoreRules) {
                    throw 'This folder is listed in .ignores.'
                }
                foreach ($profileRoot in $profileRoots) {
                    $appDataRoot = Join-Path $profileRoot 'AppData'
                    if ($fullPath.Equals($appDataRoot, [System.StringComparison]::OrdinalIgnoreCase) -or
                        $fullPath.StartsWith($appDataRoot.TrimEnd('\') + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
                        $isSafeAppDataPath = $false
                        foreach ($relativePath in $safeAppDataRelativePaths) {
                            $safeRoot = Join-Path $profileRoot $relativePath
                            if ($fullPath.Equals($safeRoot, [System.StringComparison]::OrdinalIgnoreCase) -or
                                $fullPath.StartsWith($safeRoot.TrimEnd('\') + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
                                $isSafeAppDataPath = $true
                                break
                            }
                        }
                        if (-not $isSafeAppDataPath) {
                            throw 'Only explicitly supported temporary/cache folders in AppData can be removed.'
                        }
                    }
                }
                if (-not [System.IO.Directory]::Exists($fullPath)) { throw 'Folder no longer exists.' }
                $attrs = [System.IO.File]::GetAttributes($fullPath)
                if (($attrs -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
                    throw 'Junctions and symbolic links cannot be removed by this app.'
                }
                [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory(
                    $fullPath,
                    [Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs,
                    [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin
                )
                $deleted++
                $deletedPaths.Add($fullPath)
            } catch {
                $failures.Add("$path - $($_.Exception.Message)")
            }
        }

        [pscustomobject]@{
            RecordType = 'Result'
            Deleted = $deleted
            DeletedPaths = @($deletedPaths.ToArray())
            Failures = @($failures.ToArray())
        }
    }
    try {
        $script:CleanupJob = Start-Job -ArgumentList $pathsJson, $ignoreRulesJson -ErrorAction Stop -ScriptBlock $cleanupScript
    } catch {
        $statusLabel.Text = "Unable to start cleanup: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show(
            $statusLabel.Text,
            'Cleanup could not start',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        return
    }
    $statusLabel.Text = "Moving $($paths.Count) folder(s) to the Recycle Bin..."
    $deleteButton.Text = 'Moving to Recycle Bin...'
    $deleteButton.Enabled = $false
    $scanButton.Enabled = $false
    $forceScanButton.Enabled = $false
    $thresholdInput.Enabled = $false
    $sortCombo.Enabled = $false
    $tree.Enabled = $false
    $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
    $progressBar.MarqueeAnimationSpeed = 25
    $progressBar.Visible = $true
    $timer.Start()
})

$form.Add_FormClosing({
    if ($null -ne $script:CleanupJob) {
        $args[1].Cancel = $true
        [System.Windows.Forms.MessageBox]::Show(
            'Cleanup is still processing. Wait for it to finish before closing the app.',
            'Cleanup in progress',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        return
    }
    if ($null -ne $script:ScanJob) {
        Stop-Job -Job $script:ScanJob -ErrorAction SilentlyContinue
        Remove-Job -Job $script:ScanJob -Force -ErrorAction SilentlyContinue
    }
})

$form.Add_Shown({
    try {
        $script:IgnoreRules = @(Get-IgnoreRules)
    } catch {
        $statusLabel.Text = "Unable to read .ignores: $($_.Exception.Message)"
        return
    }
    try {
        if (Import-ScanCache) { return }
    } catch {
        $statusLabel.Text = "Saved scan could not be loaded; scanning drives again: $($_.Exception.Message)"
    }
    Start-FullScan
})

[void]$form.ShowDialog()
