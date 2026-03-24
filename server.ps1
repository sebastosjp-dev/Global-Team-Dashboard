$port = 8000
$listener = New-Object System.Net.HttpListener
$serverStarted = $false

while (-not $serverStarted -and $port -le 8050) {
    try {
        $listener = New-Object System.Net.HttpListener
        $listener.Prefixes.Add("http://localhost:$port/")
        $listener.Start()
        $serverStarted = $true
    } catch {
        Write-Host "Port $port is in use or blocked, trying next port..."
        $port++
    }
}

if (-not $serverStarted) {
    Write-Host "Could not find an available port. Press any key to exit."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit
}

Write-Host "Server running at http://localhost:$port/"
Write-Host "Opening your default browser..."
Start-Process "http://localhost:$port/index.html"
Write-Host "Press Ctrl+C in this window to stop the server."

while ($listener.IsListening) {
    try {
        $context = $listener.GetContext()
        $request = $context.Request
        $response = $context.Response
        
        $path = $request.Url.LocalPath
        if ($path -eq "/") { $path = "/index.html" }
        
        # Decode path for spaces (like "global dashboard making")
        $path = [System.Uri]::UnescapeDataString($path)
        
        # Remove leading slash for joining bounds
        if ($path.StartsWith("/")) {
            $path = $path.Substring(1)
        }
        
        $scriptDir = $PSScriptRoot
        if ([string]::IsNullOrEmpty($scriptDir)) {
            $scriptDir = Get-Location
        }
        $fullPath = Join-Path $scriptDir $path
        
        if (Test-Path $fullPath -PathType Leaf) {
            try {
                # Open with ReadWrite sharing so it works EVEN IF Excel is currently open!
                $fileStream = New-Object System.IO.FileStream($fullPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
                $reader = New-Object System.IO.BinaryReader($fileStream)
                $content = $reader.ReadBytes($fileStream.Length)
                $reader.Close()
                $fileStream.Close()
                
                $response.ContentLength64 = $content.Length
                
                $ext = [System.IO.Path]::GetExtension($fullPath).ToLower()
                if ($ext -eq ".html") { $response.ContentType = "text/html; charset=utf-8" }
                elseif ($ext -eq ".js") { $response.ContentType = "application/javascript; charset=utf-8" }
                elseif ($ext -eq ".css") { $response.ContentType = "text/css; charset=utf-8" }
                elseif ($ext -eq ".xlsx") { 
                    $response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" 
                }
                
                # Disable caching for the Excel file to ensure freshest data on reload
                if ($ext -eq ".xlsx") {
                    $response.Headers.Add("Cache-Control", "no-cache, no-store, must-revalidate")
                    $response.Headers.Add("Pragma", "no-cache")
                    $response.Headers.Add("Expires", "0")
                }
                
                $response.Headers.Add("Access-Control-Allow-Origin", "*")
                
                $response.OutputStream.Write($content, 0, $content.Length)
            } catch {
                $response.StatusCode = 500
                Write-Host "Error serving file: $($_.Exception.Message)"
            }
        } else {
            $response.StatusCode = 404
        }
        $response.Close()
    } catch {
        # Catch listener aborted when closing
    }
}
