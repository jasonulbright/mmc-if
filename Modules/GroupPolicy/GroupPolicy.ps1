<#
.SYNOPSIS
    Group Policy viewer module for MMC-If (WPF).

.DESCRIPTION
    Runs `gpresult /h` to generate a per-user policy report and displays a
    plain-text summary. The full HTML report can be opened in the default
    browser. Read-only; no policy editing.
#>

function New-GroupPolicyView {
    param([hashtable]$Context)

    $setStatus = $Context.SetStatus
    $log       = $Context.Log

    $xamlPath = Join-Path $PSScriptRoot 'GroupPolicy.xaml'
    $xamlRaw  = Get-Content -LiteralPath $xamlPath -Raw
    $reader   = New-Object System.Xml.XmlNodeReader ([xml]$xamlRaw)
    $view     = [Windows.Markup.XamlReader]::Load($reader)

    $btnRun      = $view.FindName('btnRun')
    $btnOpenHtml = $view.FindName('btnOpenHtml')
    $txtOutput   = $view.FindName('txtOutput')

    $state = @{ HtmlPath = '' }

    $runGpresult = {
        $htmlPath = Join-Path $env:TEMP ("gpresult-{0}.html" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
        $state.HtmlPath = $htmlPath
        & $setStatus 'Running gpresult /h (may take 10-30 seconds)...'
        $view.Cursor = [System.Windows.Input.Cursors]::Wait
        $txtOutput.Text = '(running gpresult /h, please wait)'

        try {
            $null = & gpresult.exe /h $htmlPath /f 2>&1
            if (Test-Path -LiteralPath $htmlPath) {
                $btnOpenHtml.Visibility = [System.Windows.Visibility]::Visible
                # Also run /r for plain-text summary
                $text = & gpresult.exe /r 2>&1 | Out-String
                $txtOutput.Text = $text
                & $setStatus "gpresult complete. HTML report at $htmlPath"
                & $log "gpresult complete: $htmlPath"
            }
            else {
                $txtOutput.Text = "gpresult did not produce an HTML report. Running /r for text output:`r`n`r`n" + ((& gpresult.exe /r 2>&1) | Out-String)
                & $setStatus 'gpresult /h produced no file - text-only output shown'
            }
        }
        catch {
            $txtOutput.Text = "Error running gpresult: $($_.Exception.Message)"
            & $setStatus 'gpresult failed'
        }
        $view.Cursor = [System.Windows.Input.Cursors]::Arrow
    }.GetNewClosure()

    $btnRun.Add_Click({ & $runGpresult }.GetNewClosure())

    $btnOpenHtml.Add_Click({
        if ($state.HtmlPath -and (Test-Path -LiteralPath $state.HtmlPath)) {
            Start-Process $state.HtmlPath
        }
    }.GetNewClosure())

    return $view
}
