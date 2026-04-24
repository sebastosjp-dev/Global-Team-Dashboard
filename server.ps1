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

        $scriptDir = $PSScriptRoot
        if ([string]::IsNullOrEmpty($scriptDir)) { $scriptDir = Get-Location }

        $apiPath = $request.Url.LocalPath
        $response.Headers.Add("Access-Control-Allow-Origin", "*")

        # Helper: send JSON response
        function Send-Json($resp, $statusCode, $jsonStr) {
            $resp.StatusCode = $statusCode
            $resp.ContentType = "application/json; charset=utf-8"
            $bytes = [System.Text.Encoding]::UTF8.GetBytes($jsonStr)
            $resp.ContentLength64 = $bytes.Length
            $resp.OutputStream.Write($bytes, 0, $bytes.Length)
            $resp.Close()
        }
        # Helper: read request body
        function Read-Body($req) {
            (New-Object System.IO.StreamReader($req.InputStream, [System.Text.Encoding]::UTF8)).ReadToEnd()
        }

        # ── API: KPI Structure (Admin) ──────────────────────────────
        if ($apiPath -eq "/api/kpi/structure") {
            $year = $request.QueryString["year"]
            if ($request.HttpMethod -eq "POST") {
                $body = Read-Body $request
                [System.IO.File]::WriteAllText((Join-Path $scriptDir "kpi_structure_$year.json"), $body, [System.Text.Encoding]::UTF8)
                Write-Host "KPI structure saved: kpi_structure_$year.json"
                Send-Json $response 200 '{"ok":true}'
            } else {
                $fp = Join-Path $scriptDir "kpi_structure_$year.json"
                if (-not (Test-Path $fp)) { $fp = Join-Path $scriptDir "kpi_data_$year.json" } # legacy fallback
                if (Test-Path $fp) { Send-Json $response 200 ([System.IO.File]::ReadAllText($fp, [System.Text.Encoding]::UTF8)) }
                else { Send-Json $response 404 '{"error":"not found"}' }
            }
            continue
        }

        # ── API: KPI legacy load (backward compat) ──────────────────
        if ($apiPath -eq "/api/kpi/load" -and $request.HttpMethod -eq "GET") {
            $year = $request.QueryString["year"]
            $fp = Join-Path $scriptDir "kpi_data_$year.json"
            if (Test-Path $fp) { Send-Json $response 200 ([System.IO.File]::ReadAllText($fp, [System.Text.Encoding]::UTF8)) }
            else { Send-Json $response 404 '{"error":"not found"}' }
            continue
        }

        # ── API: KPI Achievement (per user) ─────────────────────────
        if ($apiPath -eq "/api/kpi/achievement") {
            $year = $request.QueryString["year"]
            $user = $request.QueryString["user"] -replace '[\\/:*?"<>|]', '_'  # sanitize filename
            if ($request.HttpMethod -eq "POST") {
                $body = Read-Body $request
                [System.IO.File]::WriteAllText((Join-Path $scriptDir "kpi_ach_${year}_${user}.json"), $body, [System.Text.Encoding]::UTF8)
                Write-Host "KPI achievement saved: kpi_ach_${year}_${user}.json"
                Send-Json $response 200 '{"ok":true}'
            } else {
                $fp = Join-Path $scriptDir "kpi_ach_${year}_${user}.json"
                if (Test-Path $fp) { Send-Json $response 200 ([System.IO.File]::ReadAllText($fp, [System.Text.Encoding]::UTF8)) }
                else { Send-Json $response 404 '{"error":"not found"}' }
            }
            continue
        }

        # ── API: List KPI users for a year ──────────────────────────
        if ($apiPath -eq "/api/kpi/users" -and $request.HttpMethod -eq "GET") {
            $year = $request.QueryString["year"]
            $files = Get-ChildItem $scriptDir -Filter "kpi_ach_${year}_*.json" -ErrorAction SilentlyContinue
            $users = @($files | ForEach-Object { $_.Name -replace "^kpi_ach_${year}_", "" -replace "\.json$", "" })
            Send-Json $response 200 (ConvertTo-Json $users -Compress)
            continue
        }

        # ── Static file serving ─────────────────────────────────────
        $path = $request.Url.LocalPath
        if ($path -eq "/") { $path = "/index.html" }
        $path = [System.Uri]::UnescapeDataString($path)
        if ($path.StartsWith("/")) { $path = $path.Substring(1) }
        $fullPath = Join-Path $scriptDir $path

        if (Test-Path $fullPath -PathType Leaf) {
            try {
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

                if ($ext -eq ".xlsx") {
                    $response.Headers.Add("Cache-Control", "no-cache, no-store, must-revalidate")
                    $response.Headers.Add("Pragma", "no-cache")
                    $response.Headers.Add("Expires", "0")
                }

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
