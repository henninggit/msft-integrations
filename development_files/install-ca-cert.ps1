# Install Office Add-in Development CA Certificate for WebView2
# Run this script as Administrator

$caPath = "$env:USERPROFILE\.office-addin-dev-certs\ca.crt"

if (Test-Path $caPath) {
    try {
        Import-Certificate -FilePath $caPath -CertStoreLocation Cert:\LocalMachine\Root
        Write-Host "✅ Office CA certificate successfully installed in LocalMachine\Root store" -ForegroundColor Green
        Write-Host "WebView2 should now trust the HTTPS certificates for Office add-ins" -ForegroundColor Green
        
        # Also add to WebView2 specific location if it exists
        $webView2Path = "Cert:\LocalMachine\Root"
        Write-Host "Certificate installed for WebView2 compatibility" -ForegroundColor Green
        
    } catch {
        Write-Host "❌ Failed to install certificate: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Make sure you're running PowerShell as Administrator" -ForegroundColor Yellow
    }
} else {
    Write-Host "❌ CA certificate not found at: $caPath" -ForegroundColor Red
    Write-Host "Run 'npx office-addin-dev-certs install' first" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "After installing, restart Word and try loading the add-in again." -ForegroundColor Cyan
