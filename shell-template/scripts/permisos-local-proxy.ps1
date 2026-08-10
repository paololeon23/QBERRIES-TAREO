# Proxy LOCAL seguro para Live Server (sin Node).
# Lee API_TOKEN y PERMISOS_SCRIPT_URL desde .env en la raíz del repo.
#
# Uso (PowerShell):
#   cd C:\Users\TatianaLeón\Desktop\QA-MI
#   .\shell-template\scripts\permisos-local-proxy.ps1
#
# Deja esa ventana abierta y recarga Pases en Live Server.

$ErrorActionPreference = "Stop"
$Root = Resolve-Path (Join-Path $PSScriptRoot "..\..")
$EnvFile = Join-Path $Root ".env"
$Port = 8787

function Read-DotEnv([string]$Path) {
  $map = @{}
  if (-not (Test-Path $Path)) { return $map }
  Get-Content -Path $Path -Encoding UTF8 | ForEach-Object {
    $line = $_.Trim()
    if (-not $line -or $line.StartsWith("#")) { return }
    $i = $line.IndexOf("=")
    if ($i -lt 1) { return }
    $key = $line.Substring(0, $i).Trim()
    $val = $line.Substring($i + 1).Trim()
    if (($val.StartsWith('"') -and $val.EndsWith('"')) -or ($val.StartsWith("'") -and $val.EndsWith("'"))) {
      $val = $val.Substring(1, $val.Length - 2)
    }
    $map[$key] = $val
  }
  return $map
}

$envMap = Read-DotEnv $EnvFile
$Token = ""
$ScriptUrl = ""
if ($envMap.ContainsKey("API_TOKEN") -and $envMap["API_TOKEN"]) { $Token = [string]$envMap["API_TOKEN"] }
elseif ($envMap.ContainsKey("PERMISOS_API_TOKEN") -and $envMap["PERMISOS_API_TOKEN"]) { $Token = [string]$envMap["PERMISOS_API_TOKEN"] }
if ($envMap.ContainsKey("PERMISOS_SCRIPT_URL") -and $envMap["PERMISOS_SCRIPT_URL"]) {
  $ScriptUrl = [string]$envMap["PERMISOS_SCRIPT_URL"]
}

if (-not $ScriptUrl) {
  Write-Host "Falta PERMISOS_SCRIPT_URL en .env (solo ahí; no hardcodear en el repo)."
  exit 1
}

$Allowed = @("ping", "listarPermisos", "obtenerPermiso", "existePaseHoy")

$listener = New-Object System.Net.HttpListener
$prefix = "http://127.0.0.1:$Port/"
$listener.Prefixes.Add($prefix)

try {
  $listener.Start()
} catch {
  Write-Host "No se pudo abrir $prefix"
  Write-Host $_.Exception.Message
  Write-Host "Cierra lo que use el puerto $Port o cambia el puerto en pases-api.js"
  exit 1
}

Write-Host "[permisos-local-proxy] listening $prefix"
Write-Host "[permisos-local-proxy] token cargado: $([bool]$Token) | script ok: $([bool]$ScriptUrl)"
Write-Host "Deja esta ventana abierta. Recarga Pases en Live Server."

while ($listener.IsListening) {
  $ctx = $listener.GetContext()
  $req = $ctx.Request
  $res = $ctx.Response

  $origin = $req.Headers["Origin"]
  if (-not $origin) { $origin = "http://127.0.0.1:5500" }
  $res.Headers.Add("Access-Control-Allow-Origin", $origin)
  $res.Headers.Add("Access-Control-Allow-Methods", "GET, OPTIONS")
  $res.Headers.Add("Access-Control-Allow-Headers", "Content-Type, Accept")
  $res.Headers.Add("Cache-Control", "no-store")
  $res.Headers.Add("Vary", "Origin")

  if ($req.HttpMethod -eq "OPTIONS") {
    $res.StatusCode = 204
    $res.Close()
    continue
  }

  $path = $req.Url.AbsolutePath.TrimEnd("/")
  if ($path -ne "/api/permisos") {
    $bytes = [Text.Encoding]::UTF8.GetBytes('{"ok":false,"message":"Usa /api/permisos"}')
    $res.StatusCode = 404
    $res.ContentType = "application/json; charset=utf-8"
    $res.OutputStream.Write($bytes, 0, $bytes.Length)
    $res.Close()
    continue
  }

  if ($req.HttpMethod -ne "GET") {
    $bytes = [Text.Encoding]::UTF8.GetBytes('{"ok":false,"message":"Solo GET"}')
    $res.StatusCode = 405
    $res.ContentType = "application/json; charset=utf-8"
    $res.OutputStream.Write($bytes, 0, $bytes.Length)
    $res.Close()
    continue
  }

  if (-not $Token) {
    $bytes = [Text.Encoding]::UTF8.GetBytes('{"ok":false,"code":"CONFIG","message":"Falta API_TOKEN en .env"}')
    $res.StatusCode = 500
    $res.ContentType = "application/json; charset=utf-8"
    $res.OutputStream.Write($bytes, 0, $bytes.Length)
    $res.Close()
    continue
  }

  $params = @{}
  if ($req.Url.Query.Length -gt 1) {
    $req.Url.Query.TrimStart("?") -split "&" | ForEach-Object {
      if (-not $_) { return }
      $kv = $_.Split("=", 2)
      $k = [Uri]::UnescapeDataString($kv[0])
      $v = if ($kv.Count -gt 1) { [Uri]::UnescapeDataString($kv[1].Replace("+", " ")) } else { "" }
      $params[$k] = $v
    }
  }

  $action = if ($params.ContainsKey("action") -and $params["action"]) { $params["action"] } else { "ping" }
  if ($Allowed -notcontains $action) {
    $bytes = [Text.Encoding]::UTF8.GetBytes('{"ok":false,"code":"FORBIDDEN","message":"Accion no permitida"}')
    $res.StatusCode = 403
    $res.ContentType = "application/json; charset=utf-8"
    $res.OutputStream.Write($bytes, 0, $bytes.Length)
    $res.Close()
    continue
  }

  $parts = New-Object System.Collections.Generic.List[string]
  foreach ($key in $params.Keys) {
    if ($key -match '^(?i)(callback|token|apitoken|api_token)$') { continue }
    if (-not $params[$key]) { continue }
    $parts.Add(("{0}={1}" -f [Uri]::EscapeDataString($key), [Uri]::EscapeDataString([string]$params[$key])))
  }
  $parts.Add(("token={0}" -f [Uri]::EscapeDataString($Token)))
  $q = ($parts -join "&")

  try {
    $upstream = Invoke-WebRequest -Uri ($ScriptUrl + "?" + $q) -Method GET -UseBasicParsing -TimeoutSec 60
    $bodyText = $upstream.Content
    $bytes = [Text.Encoding]::UTF8.GetBytes($bodyText)
    $res.StatusCode = 200
    $res.ContentType = "application/json; charset=utf-8"
    $res.OutputStream.Write($bytes, 0, $bytes.Length)
  } catch {
    $msg = $_.Exception.Message -replace '"', "'"
    $json = "{{""ok"":false,""code"":""PROXY"",""message"":""{0}""}}" -f $msg
    $bytes = [Text.Encoding]::UTF8.GetBytes($json)
    $res.StatusCode = 502
    $res.ContentType = "application/json; charset=utf-8"
    $res.OutputStream.Write($bytes, 0, $bytes.Length)
  }
  $res.Close()
}
