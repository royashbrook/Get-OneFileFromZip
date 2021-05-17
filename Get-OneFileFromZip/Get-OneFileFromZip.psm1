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
Export-ModuleMember -Function Get-OneFileFromZip