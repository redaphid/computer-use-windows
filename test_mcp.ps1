$baseUrl = 'http://127.0.0.1:8080/mcp'

$headers = @{
    'Accept' = 'application/json, text/event-stream'
}

$initBody = @{
    jsonrpc = "2.0"
    method = "initialize"
    id = 1
    params = @{
        protocolVersion = "2024-11-05"
        capabilities = @{}
        clientInfo = @{
            name = "test-client"
            version = "1.0.0"
        }
    }
} | ConvertTo-Json -Depth 5

Write-Host "=== Testing MCP Server ===" -ForegroundColor Cyan
Write-Host ""
Write-Host "1. Initialize Request:" -ForegroundColor Yellow

$sessionId = $null
try {
    $response = Invoke-WebRequest -Uri $baseUrl -Method POST -ContentType 'application/json' -Headers $headers -Body $initBody
    Write-Host "Status: $($response.StatusCode)" -ForegroundColor Green

    # Get session ID from header
    $sessionId = $response.Headers['Mcp-Session-Id']
    if ($sessionId) {
        Write-Host "Session ID: $sessionId" -ForegroundColor Green
    }

    # Parse SSE response
    $content = $response.Content
    if ($content -match 'data: (.+)') {
        $jsonData = $Matches[1] | ConvertFrom-Json
        $jsonData | ConvertTo-Json -Depth 10
    }
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}

if ($sessionId) {
    $headers['Mcp-Session-Id'] = $sessionId
}

Write-Host ""
Write-Host "2. List Tools Request:" -ForegroundColor Yellow

$listToolsBody = @{
    jsonrpc = "2.0"
    method = "tools/list"
    id = 2
    params = @{}
} | ConvertTo-Json -Depth 5

try {
    $response = Invoke-WebRequest -Uri $baseUrl -Method POST -ContentType 'application/json' -Headers $headers -Body $listToolsBody
    $content = $response.Content
    if ($content -match 'data: (.+)') {
        $jsonData = $Matches[1] | ConvertFrom-Json
        Write-Host "Available Tools:" -ForegroundColor Green
        foreach ($tool in $jsonData.result.tools) {
            Write-Host "  - $($tool.name): $($tool.description)" -ForegroundColor White
        }
    }
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Response: $($_.ErrorDetails.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "3. Call get_screen_size Tool:" -ForegroundColor Yellow

$callToolBody = @{
    jsonrpc = "2.0"
    method = "tools/call"
    id = 3
    params = @{
        name = "get_screen_size"
        arguments = @{}
    }
} | ConvertTo-Json -Depth 5

try {
    $response = Invoke-WebRequest -Uri $baseUrl -Method POST -ContentType 'application/json' -Headers $headers -Body $callToolBody
    $content = $response.Content
    if ($content -match 'data: (.+)') {
        $jsonData = $Matches[1] | ConvertFrom-Json
        Write-Host "Screen Size Result:" -ForegroundColor Green
        $jsonData.result | ConvertTo-Json -Depth 10
    }
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Response: $($_.ErrorDetails.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "4. Call get_mouse_position Tool:" -ForegroundColor Yellow

$callToolBody = @{
    jsonrpc = "2.0"
    method = "tools/call"
    id = 4
    params = @{
        name = "get_mouse_position"
        arguments = @{}
    }
} | ConvertTo-Json -Depth 5

try {
    $response = Invoke-WebRequest -Uri $baseUrl -Method POST -ContentType 'application/json' -Headers $headers -Body $callToolBody
    $content = $response.Content
    if ($content -match 'data: (.+)') {
        $jsonData = $Matches[1] | ConvertFrom-Json
        Write-Host "Mouse Position Result:" -ForegroundColor Green
        $jsonData.result | ConvertTo-Json -Depth 10
    }
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Response: $($_.ErrorDetails.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "=== Test Complete ===" -ForegroundColor Cyan
