# Autodoc v1.0
function Show-CollectorPicker {
    <#
        Shows the Autodoc window (layout in Entra\Config\CollectorPicker.xaml).
        The window STAYS OPEN after Run so several actions can be performed in
        one session; close it with the Close button or the window X.

        The active tab decides the action passed to -OnRun:
          Download Data       -> @{ Action='Collect'; Collectors=@(); Formats=@() }
          Generate Documents  -> @{ Action='ReportDocs'; Formats=@(ticked); SourceFolder=... }

        NOTE: the work runs on the UI thread, so the window is unresponsive while
        a job is running (progress is written to the console). It becomes usable
        again when the run finishes.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Registry,
        [Parameter(Mandatory)][scriptblock]$OnRun,
        [string]$ExportsRoot,
        [string]$XamlPath
    )

    Add-Type -AssemblyName PresentationFramework

    # The window is application-level rather than Entra-specific, so the layout
    # lives in the project root's Config folder: three levels up from
    # Entra\Functions\GUI.
    if (-not $XamlPath) {
        $projectRoot = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
        $XamlPath    = Join-Path $projectRoot 'Config\CollectorPicker.xaml'
    }
    if (-not (Test-Path $XamlPath)) { throw "XAML layout not found: $XamlPath" }

    [xml]$xaml = Get-Content -Path $XamlPath -Raw
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $window = [Windows.Markup.XamlReader]::Load($reader)

    $tabs   = $window.FindName('ActionTabs')
    $list   = $window.FindName('CollectorList')
    $status = $window.FindName('StatusText')

    # One checkbox per collector, driven by the registry
    foreach ($def in $Registry) {
        $cb = New-Object System.Windows.Controls.CheckBox
        $cb.IsChecked = $true
        $cb.Tag       = $def.Key
        $cb.Margin    = [System.Windows.Thickness]::new(0, 6, 0, 6)

        $stack = New-Object System.Windows.Controls.StackPanel
        $title = New-Object System.Windows.Controls.TextBlock
        $title.Text       = $def.Name
        $title.FontWeight = [System.Windows.FontWeights]::SemiBold
        $desc = New-Object System.Windows.Controls.TextBlock
        $desc.Text        = $def.Description
        $desc.Foreground  = [System.Windows.Media.Brushes]::Gray
        $desc.TextWrapping = [System.Windows.TextWrapping]::Wrap
        $stack.Children.Add($title) | Out-Null
        $stack.Children.Add($desc)  | Out-Null
        $cb.Content = $stack

        $list.Children.Add($cb) | Out-Null
    }

    # --- report source folder: newest export preselected, Browse for anything else ---
    $sourceCombo = $window.FindName('SourceFolders')
    $populateSource = {
        $prev = "$($sourceCombo.SelectedItem)"
        $sourceCombo.Items.Clear()
        if ($ExportsRoot -and (Test-Path $ExportsRoot)) {
            foreach ($folder in (Get-ChildItem -Path $ExportsRoot -Directory | Sort-Object Name -Descending)) {
                $null = $sourceCombo.Items.Add($folder.FullName)
            }
        }
        if ($sourceCombo.Items.Count -gt 0) {
            if ($prev -and $sourceCombo.Items.Contains($prev)) { $sourceCombo.SelectedItem = $prev }   # keep current pick if still there
            else { $sourceCombo.SelectedIndex = 0 }   # otherwise latest
        }
        else { $null = $sourceCombo.Items.Add('(no exports found - run a collection first)'); $sourceCombo.SelectedIndex = 0 }
    }
    & $populateSource

    $window.FindName('BtnReloadSource').Add_Click({ & $populateSource })

    $window.FindName('BtnBrowse').Add_Click({
        Add-Type -AssemblyName System.Windows.Forms
        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.Description = 'Select the export folder to build the report from'
        if ($ExportsRoot -and (Test-Path $ExportsRoot)) { $dialog.SelectedPath = $ExportsRoot }
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $picked = $dialog.SelectedPath
            if (-not $sourceCombo.Items.Contains($picked)) { $null = $sourceCombo.Items.Insert(0, $picked) }
            $sourceCombo.SelectedItem = $picked
        }
    })

    # Report source only applies to the report tabs (Excel/SVG/HTML/Word), not to
    # Download Data (which collects fresh data). Hide the panel on that tab.
    $actionTabs        = $window.FindName('ActionTabs')
    $reportSourcePanel = $window.FindName('ReportSourcePanel')
    $updateSourceVisibility = {
        $header = "$($actionTabs.SelectedItem.Header)"
        if ($header -eq 'Download Data') {
            $reportSourcePanel.Visibility = [System.Windows.Visibility]::Collapsed
        } else {
            $reportSourcePanel.Visibility = [System.Windows.Visibility]::Visible
        }
    }
    $actionTabs.Add_SelectionChanged($updateSourceVisibility)
    & $updateSourceVisibility   # set correct state on open (Download Data is first)

    $exportCsv   = $window.FindName('ExportCsv')
    $exportJson  = $window.FindName('ExportJson')
    $exportExcel = $window.FindName('ExportExcel')

    # document-format checkboxes on the Generate Documents tab
    $docExcel = $window.FindName('DocExcel')
    $docSvg   = $window.FindName('DocSvg')
    $docWord  = $window.FindName('DocWord')
    $docHtml  = $window.FindName('DocHtml')

    # Populate the Word language list from the folders under wordTemplates, so
    # adding a language folder (en, sv, es, ...) makes it selectable here.
    $wordLang = $window.FindName('WordLanguage')
    $tmplRoot = Join-Path (Split-Path (Split-Path $PSScriptRoot)) 'files\cache\wordTemplates'
    if (Test-Path $tmplRoot) {
        foreach ($d in (Get-ChildItem -Path $tmplRoot -Directory | Sort-Object Name)) { $null = $wordLang.Items.Add($d.Name) }
    }
    if ($wordLang.Items.Contains('en')) { $wordLang.SelectedItem = 'en' }
    elseif ($wordLang.Items.Count -gt 0) { $wordLang.SelectedIndex = 0 }

    # optional logo pickers for the Word cover page.
    # Uses the WPF-native dialog (no WinForms load) and copies the picked file into
    # files\cache\wordTemplates\template\img so the Word builder reads it from there.
    $wordCompanyLogo = $window.FindName('WordCompanyLogo')
    $wordClientLogo  = $window.FindName('WordClientLogo')
    $logoImgDir = Join-Path (Split-Path (Split-Path $PSScriptRoot)) 'files\cache\wordTemplates\template\img'
    $pickLogo = {
        param($textbox, $baseName)
        $ofd = New-Object Microsoft.Win32.OpenFileDialog
        $ofd.Filter = 'Images (*.png;*.jpg;*.jpeg)|*.png;*.jpg;*.jpeg|All files (*.*)|*.*'
        if ($ofd.ShowDialog() -eq $true) {
            try {
                if (-not (Test-Path $logoImgDir)) { New-Item -Path $logoImgDir -ItemType Directory -Force | Out-Null }
                Get-ChildItem -Path $logoImgDir -Filter "$baseName.*" -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
                $dest = Join-Path $logoImgDir ($baseName + [System.IO.Path]::GetExtension($ofd.FileName))
                Copy-Item -LiteralPath $ofd.FileName -Destination $dest -Force
                $textbox.Text = $dest
            }
            catch { $textbox.Text = $ofd.FileName }
        }
    }
    $window.FindName('BtnCompanyLogo').Add_Click({ & $pickLogo $wordCompanyLogo 'companylogo' })
    $window.FindName('BtnClientLogo').Add_Click({ & $pickLogo $wordClientLogo 'clientlogo' })

    $btnRun      = $window.FindName('BtnRun')

    $window.FindName('BtnAll').Add_Click( { foreach ($c in $list.Children) { $c.IsChecked = $true } } )
    $window.FindName('BtnNone').Add_Click({ foreach ($c in $list.Children) { $c.IsChecked = $false } })

    # Disconnect from Microsoft Graph so the next collection prompts a fresh
    # sign-in - lets you run against a different tenant.
    $window.FindName('BtnDisconnect').Add_Click({
        try {
            Disconnect-MgGraph -ErrorAction Stop | Out-Null
            $status.Text = 'Signed out of Microsoft Graph. The next Run will prompt sign-in - pick the tenant there.'
        }
        catch {
            $status.Text = 'Not currently signed in to Microsoft Graph.'
        }
    })

    # Run: build the request from the active tab, invoke the callback, keep the window open
    $btnRun.Add_Click({
        $header  = [string]$tabs.SelectedItem.Header
        $chosen  = [string]$sourceCombo.SelectedItem
        if ($chosen -like '(no exports*') { $chosen = $null }
        $request = switch ($header) {
            'Generate Documents' {
                $formats = @()
                if ($docExcel.IsChecked) { $formats += 'Excel' }
                if ($docSvg.IsChecked)   { $formats += 'Svg' }
                if ($docWord.IsChecked)  { $formats += 'Word' }
                if ($docHtml.IsChecked)  { $formats += 'Html' }
                [pscustomobject]@{ Action = 'ReportDocs'; Formats = $formats; SourceFolder = $chosen; WordLanguage = [string]$wordLang.SelectedItem; WordCompanyLogo = [string]$wordCompanyLogo.Text; WordClientLogo = [string]$wordClientLogo.Text }
            }
            default {
                $selectedCollectors = @(foreach ($c in $list.Children) { if ($c.IsChecked) { [string]$c.Tag } })
                $selectedFormats = @()
                if ($exportCsv.IsChecked)   { $selectedFormats += 'Csv' }
                if ($exportJson.IsChecked)  { $selectedFormats += 'Json' }
                if ($exportExcel.IsChecked) { $selectedFormats += 'Excel' }
                [pscustomobject]@{ Action = 'Collect'; Collectors = $selectedCollectors; Formats = $selectedFormats }
            }
        }

        $btnRun.IsEnabled = $false
        $status.Text      = "Running: $header - see the console for progress..."
        # Let WPF paint the disabled state before the (blocking) work starts
        $window.Dispatcher.Invoke([action]{}, [Windows.Threading.DispatcherPriority]::Render)

        try   { & $OnRun $request; $status.Text = "Finished: $header at $(Get-Date -Format 'HH:mm:ss')" }
        catch { $status.Text = "Failed: $($_.Exception.Message)" ; Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red }
        finally { $btnRun.IsEnabled = $true }
    })

    $null = $window.ShowDialog()
}
