import subprocess

verify_ps = """
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $filePath = 'D:\\audit-log-system\\BAO_CAO_DO_AN_CUOI_KY_DAC_TA_HE_THONG_AUDIT_LOG_NHOM_13.docx'
    $doc = $word.Documents.Open($filePath)
    
    # Update all fields in document (page numbers, etc.)
    foreach ($story in $doc.StoryRanges) {
        $story.Fields.Update()
    }
    
    $pages = $doc.ComputeStatistics(2) # 2 = wdStatisticPages
    $words = $doc.ComputeStatistics(0) # 0 = wdStatisticWords
    $paragraphs = $doc.Paragraphs.Count
    $tables = $doc.Tables.Count
    $sections = $doc.Sections.Count
    
    Write-Host "DOCX Verification Results:"
    Write-Host "--------------------------"
    Write-Host "Total Pages:       $pages"
    Write-Host "Total Words:       $words"
    Write-Host "Total Paragraphs:  $paragraphs"
    Write-Host "Total Tables:      $tables"
    Write-Host "Total Sections:    $sections"
    
    $doc.Save()
    $doc.Close()
    Write-Host "File saved and closed cleanly!"
} catch {
    Write-Error $_
} finally {
    $word.Quit()
}
"""

with open("verify_doc.ps1", "w", encoding="utf-8") as f:
    f.write(verify_ps)

res = subprocess.run(["powershell", "-ExecutionPolicy", "Bypass", "-File", "verify_doc.ps1"], capture_output=True, text=True, encoding="utf-8")
print("STDOUT:")
print(res.stdout)
if res.stderr:
    print("STDERR:")
    print(res.stderr)
