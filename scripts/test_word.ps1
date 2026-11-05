try {
    $word = New-Object -ComObject Word.Application
    if ($word) {
        Write-Output "WORD_OK"
        $word.Quit()
    }
} catch {
    Write-Output "WORD_FAILED: $_"
}
