function Get-OneFileFromZip($Path,$File){
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    $entry = $zip.entries | Where-Object{$_.Name -eq $File}
    $stream = $entry.Open()
    $reader = New-Object System.IO.StreamReader($stream)
    $reader.ReadToEnd()
    $null=$reader.Dispose()
    $null=$stream.Dispose()
    $null=$zip.Dispose()
}
function Get-SheetByName($Path,$SheetName){

    # see the following link for details on spreadsheetml structure
    # https://docs.microsoft.com/en-us/office/open-xml/structure-of-a-spreadsheetml-document
    
    # get internal files that hold sheet mappings
    [xml]$sheets = Get-OneFileFromZip $Path 'workbook.xml'
    [xml]$rels = Get-OneFileFromZip $Path 'workbook.xml.rels'
    #resolve internal file name for sheet by name
    $rid = ($sheets.workbook.sheets.sheet | Where-Object name -eq $SheetName).id

    $sheetfile = split-path -leaf ($rels.Relationships.Relationship | Where-Object id -eq $rid).Target
    Get-OneFileFromZip $Path $sheetfile
 }
Export-ModuleMember -Function Get-OneFileFromZip,Get-SheetByName