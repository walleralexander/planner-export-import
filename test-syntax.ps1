# Minimal test for try-catch structure
foreach ($item in 1..2) {
    try {
        Write-Host "Item: $item"
        
        # Inner block 1
        if ($true) {
            try {
                Write-Host "Inner try 1"
            }
            catch {
                Write-Host "Inner catch 1"
            }
        }
        
        # Logic after inner block
        $array = @()
        
        # Inner block 2
        if ($true) {
            try {
                Write-Host "Inner try 2"
            }
            catch {
                Write-Host "Inner catch 2"
            }
        }
    }
    catch {
        Write-Host "Outer catch"
    }
}
