# Servidor HTTP simples para a rede local
$porta = 8080
$pasta = $PSScriptRoot

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://+:$porta/")

try {
    $listener.Start()
} catch {
    Write-Host ""
    Write-Host "ERRO: Execute este script como Administrador." -ForegroundColor Red
    Write-Host "Clique com botao direito no servidor.bat e escolha 'Executar como administrador'" -ForegroundColor Yellow
    Read-Host "Pressione Enter para fechar"
    exit
}

# Pega os IPs da maquina
$ips = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.IPAddress -notlike "127.*" -and $_.IPAddress -notlike "169.*" }) | Select-Object -ExpandProperty IPAddress

Write-Host ""
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host "  SERVIDOR RODANDO NA PORTA $porta" -ForegroundColor Green
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Acesse no seu computador:" -ForegroundColor White
Write-Host "  http://localhost:$porta" -ForegroundColor Yellow
Write-Host ""
Write-Host "Acesse de outros computadores da rede:" -ForegroundColor White
foreach ($ip in $ips) {
    Write-Host "  http://${ip}:$porta" -ForegroundColor Yellow
}
Write-Host ""
Write-Host "Pressione Ctrl+C para parar o servidor." -ForegroundColor Gray
Write-Host ""

$tiposMime = @{
    ".html" = "text/html; charset=utf-8"
    ".css"  = "text/css"
    ".js"   = "application/javascript"
    ".png"  = "image/png"
    ".jpg"  = "image/jpeg"
    ".ico"  = "image/x-icon"
    ".json" = "application/json"
}

while ($listener.IsListening) {
    try {
        $ctx = $listener.GetContext()
        $req = $ctx.Request
        $res = $ctx.Response

        $caminho = $req.Url.LocalPath
        if ($caminho -eq "/") { $caminho = "/index.html" }

        $arquivo = Join-Path $pasta $caminho.TrimStart("/")

        if (Test-Path $arquivo -PathType Leaf) {
            $ext = [System.IO.Path]::GetExtension($arquivo)
            $mime = if ($tiposMime.ContainsKey($ext)) { $tiposMime[$ext] } else { "application/octet-stream" }
            $bytes = [System.IO.File]::ReadAllBytes($arquivo)
            $res.ContentType = $mime
            $res.ContentLength64 = $bytes.Length
            $res.OutputStream.Write($bytes, 0, $bytes.Length)
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] 200 $caminho" -ForegroundColor Green
        } else {
            $msg = [System.Text.Encoding]::UTF8.GetBytes("Arquivo nao encontrado: $caminho")
            $res.StatusCode = 404
            $res.ContentLength64 = $msg.Length
            $res.OutputStream.Write($msg, 0, $msg.Length)
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] 404 $caminho" -ForegroundColor Red
        }

        $res.OutputStream.Close()
    } catch {
        if ($listener.IsListening) {
            Write-Host "Erro: $_" -ForegroundColor Red
        }
    }
}
