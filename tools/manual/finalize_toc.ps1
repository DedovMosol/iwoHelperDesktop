# Materialize the table of contents INSIDE the .docx.
#
# python-docx can only write a TOC *field* - the entries and their page numbers
# are computed by Word, nobody else. to_pdf.ps1 updates the fields on the way to
# PDF, but it opens the document read-only, so the .docx that ships in the repo
# kept an empty contents page and a note asking the reader to press F9 himself.
# This step opens the document read-write, updates the fields, repaginates and
# saves the result back, so the shipped manual has a real table of contents.
#
# The path must be ASCII on purpose: Windows PowerShell 5.1 mangles non-ASCII
# text, so the caller hands over an ASCII-named copy and renames it afterwards.
param([Parameter(Mandatory = $true)][string]$Path)

$ErrorActionPreference = 'Stop'
if (-not (Test-Path -LiteralPath $Path)) { Write-Error "no such file: $Path"; exit 2 }

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$code = 1
try {
    $doc = $word.Documents.Open($Path, $false, $false)  # ConfirmConversions=false, ReadOnly=false
    $doc.Fields.Update() | Out-Null
    if ($doc.TablesOfContents.Count -lt 1) { throw 'the document carries no table-of-contents field' }
    $doc.TablesOfContents.Item(1).Update() | Out-Null
    # The entries must not inherit the body's first-line indent. Word rebuilds the
    # toc styles from its own built-ins (they carry w:autoRedefine, so a first-line
    # indent set on the style beforehand is thrown away) and bases them on Normal,
    # which starts every paragraph 1.25 cm in - including every line of contents.
    $doc.TablesOfContents.Item(1).Range.ParagraphFormat.FirstLineIndent = 0
    $doc.Repaginate()
    # An empty field would save just as happily as a filled one, so assert on the
    # result: this step exists precisely because nobody noticed an empty contents.
    $entries = $doc.TablesOfContents.Item(1).Range.Paragraphs.Count
    Write-Output ("pages=" + $doc.ComputeStatistics(2))
    Write-Output ("toc_entries=" + $entries)
    if ($entries -lt 10) { throw "the table of contents came out with $entries entries" }
    $doc.Save()
    $doc.Close($false)
    Write-Output 'ok'
    $code = 0
} finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
exit $code
