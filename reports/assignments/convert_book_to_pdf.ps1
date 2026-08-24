# Converts the book docx to PDF via Word COM automation, updating the TOC
# field first. python-docx only inserts the TOC field code, not its computed
# result - a plain SaveAs never triggers Word to calculate it, so the Table
# of Contents exports as a blank page under its own heading unless the field
# (and pagination, which the field's own line count then shifts) is updated
# before saving. Usage: powershell -File convert_book_to_pdf.ps1 <docx> <pdf>
#
# SaveAs to PDF right after Repaginate()/Fields.Update() intermittently
# throws a transient COMException (Word still busy internally) - PowerShell
# treats this as non-terminating by default and keeps running, which used to
# print "Saved" even when SaveAs never actually wrote the file, silently
# leaving a stale PDF from a previous run in place. This version makes that
# failure fatal, retries a few times, and verifies the output file's mtime
# actually advanced past the moment this run started before declaring success.
param(
    [Parameter(Mandatory=$true)][string]$InPath,
    [Parameter(Mandatory=$true)][string]$OutPath
)
$ErrorActionPreference = "Stop"
$runStart = Get-Date
$maxAttempts = 3

for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
    Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 1
    $word = $null
    $doc = $null
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $doc = $word.Documents.Open((Resolve-Path $InPath).Path)
        foreach ($toc in $doc.TablesOfContents) { $toc.Update() }
        $doc.Fields.Update() | Out-Null
        foreach ($toc in $doc.TablesOfContents) { $toc.Update() }
        $doc.Repaginate()
        $doc.SaveAs([ref]$OutPath, [ref]17)
        $doc.Close([ref]$false)
        $word.Quit()

        $outInfo = Get-Item $OutPath -ErrorAction Stop
        if ($outInfo.LastWriteTime -lt $runStart) {
            throw "SaveAs returned without error but $OutPath was not actually updated (mtime $($outInfo.LastWriteTime) predates this run)."
        }
        Write-Output "Saved $OutPath"
        exit 0
    } catch {
        Write-Warning "Attempt $attempt/$maxAttempts failed: $($_.Exception.Message)"
        if ($doc) { try { $doc.Close([ref]$false) } catch {} }
        if ($word) { try { $word.Quit() } catch {} }
        Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
        if ($attempt -eq $maxAttempts) {
            Write-Error "Failed to convert $InPath to $OutPath after $maxAttempts attempts."
            exit 1
        }
        Start-Sleep -Seconds 2
    }
}
