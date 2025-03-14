# Funciones

function Show-Menu {
    param (
        [string[]]$Options,
        [string]$Prompt,
        [switch]$MultiSelect,
        [int]$DefaultIndex = 0
    )

    $selectedIndex = $DefaultIndex
    $selectedOptions = @($false) * $Options.Count

    while ($true) {
        Clear-Host
        Write-Host $Prompt -ForegroundColor Cyan
        Write-Host "Use las flechas para navegar, Enter para seleccionar" -ForegroundColor Yellow
        if ($MultiSelect) {
            Write-Host "Presione Espacio para marcar/desmarcar multiples opciones" -ForegroundColor Yellow
        }
        Write-Host ""

        for ($i = 0; $i -lt $Options.Count; $i++) {
            if ($i -eq $selectedIndex) {
                $prefix = if ($MultiSelect) {
                    if ($selectedOptions[$i]) { "-> [X]" } else { "-> [ ]" }
                } else {
                    "->"
                }
                Write-Host "$prefix $($Options[$i])" -ForegroundColor Green
            } else {
                $prefix = if ($MultiSelect) {
                    if ($selectedOptions[$i]) { "   [X]" } else { "   [ ]" }
                } else {
                    "  "
                }
                Write-Host "$prefix $($Options[$i])" -ForegroundColor White
            }
        }

        $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

        switch ($key.VirtualKeyCode) {
            38 { $selectedIndex = ($selectedIndex - 1 + $Options.Count) % $Options.Count } # Flecha arriba
            40 { $selectedIndex = ($selectedIndex + 1) % $Options.Count } # Flecha abajo
            32 { if ($MultiSelect) { $selectedOptions[$selectedIndex] = -not $selectedOptions[$selectedIndex] } } # Espacio
            13 { # Enter
                if ($MultiSelect) {
                    return $Options.Where({ $selectedOptions[$Options.IndexOf($_)] })
                } else {
                    return $Options[$selectedIndex]
                }
            }
        }
    }
}

function EnsureSetupExeExists {
    $setupUrl = 'https://github.com/astronomyc/astrofficePs1/raw/main/setup.exe'
    $setupPath = "$env:TEMP\Setup.exe"

    if (-not (Test-Path $setupPath)) {
        Write-Host "Descargando setup.exe..." -ForegroundColor Cyan
        (New-Object System.Net.WebClient).DownloadFile($setupUrl, $setupPath)
        Write-Host "Descarga completada." -ForegroundColor Green
    }

    return $setupPath
}

function InstallOffice {
    $officeProducts = @("Word", "Excel", "PowerPoint", "Access", "Publisher", "OneNote", "Outlook", "Teams", "OneDrive", "Lync", "Project", "Visio")
    $selectedProducts = Show-Menu -Options $officeProducts -Prompt "Seleccione los productos de Office a instalar:" -MultiSelect

    $arch = Show-Menu -Options @("x64", "x32") -Prompt "Seleccione la arquitectura:" -DefaultIndex 0
    $arch = if ($arch -eq "x64") { 64 } else { 32 }

    $productIdOptions = @(
        "ProPlus2019Volume",
        "Standard2019Volume",
        "ProPlus2021Volume",
        "Standard2021Volume",
        "ProPlus2024Volume",
        "Standard2024Volume",
        "O365ProPlusRetail"
    )
    $productId = Show-Menu -Options $productIdOptions -Prompt "Seleccione el ID del producto:" -DefaultIndex 4

    $channel = switch -Wildcard ($productId) {
        "*2019*" { "PerpetualVL2019" }
        "*2021*" { "PerpetualVL2021" }
        "*2024*" { "PerpetualVL2024" }
        "O365*" { "Current" }
        default { "PerpetualVL2021" }
    }

    $languageOptions = @("Espanol", "English", "Francais", "Deutsch")
    $language = Show-Menu -Options $languageOptions -Prompt "Seleccione el idioma:" -DefaultIndex 0
    $languageId = switch ($language) {
        "Espanol" { "es-es" }
        "English" { "en-us" }
        "Francais" { "fr-fr" }
        "Deutsch" { "de-de" }
        default { "en-us" }
    }

    # Generar la configuracion XML
    $xmlContent = @"
<Configuration>
    <Add OfficeClientEdition="$arch" Channel="$channel">
        <Product ID="$productId">
            <Language ID="$languageId" />
"@

    $excludedApps = $officeProducts | Where-Object { $_ -notin $selectedProducts }
    foreach ($app in $excludedApps) {
        if ($app -notin @("Project", "Visio")) {
            $xmlContent += "            <ExcludeApp ID=""$app"" />`n"
        }
    }

    $xmlContent += @"
        </Product>
"@

    if ($selectedProducts -contains "Project") {
        $xmlContent += @"
        <Product ID="ProjectProRetail">
            <Language ID="$languageId" />
        </Product>
"@
    }

    if ($selectedProducts -contains "Visio") {
        $xmlContent += @"
        <Product ID="VisioProRetail">
            <Language ID="$languageId" />
        </Product>
"@
    }

    $xmlContent += @"
    </Add>
    <Updates Enabled="False" />
</Configuration>
"@

    $configPath = "$env:TEMP\config.xml"
    $xmlContent | Out-File -FilePath $configPath -Encoding utf8

    $setupPath = EnsureSetupExeExists

    Write-Host "Iniciando la instalacion de Office..." -ForegroundColor Cyan
    Start-Process -FilePath $setupPath -ArgumentList "/configure `"$configPath`"" -Wait
    Write-Host "Instalacion completada." -ForegroundColor Green
}

function ActivateOffice {
    $url = "https://raw.githubusercontent.com/astronomyc/astrofficePs1/refs/heads/main/activar.txt"
    $destination = "$env:TEMP\activate.cmd"
    (New-Object System.Net.WebClient).DownloadFile($url, $destination)

    & $destination /Ohook
    Write-Host "Activacion completa" -ForegroundColor Green
}

function UninstallOffice {
    Write-Host "Iniciando la desinstalacion de Office..." -ForegroundColor Cyan

    $setupPath = EnsureSetupExeExists

    # Crear el XML de configuracion para desinstalar
    $uninstallXml = @"
<Configuration>
    <Remove All="True" />
    <Display Level="None" AcceptEULA="TRUE" />
</Configuration>
"@
    $uninstallConfigPath = "$env:TEMP\uninstall_config.xml"
    $uninstallXml | Out-File -FilePath $uninstallConfigPath -Encoding utf8

    # Ejecutar la desinstalacion
    Start-Process -FilePath $setupPath -ArgumentList "/configure `"$uninstallConfigPath`"" -Wait

    Write-Host "Desinstalacion completada." -ForegroundColor Green
}

# Menu principal
while ($true) {
    $mainOption = Show-Menu -Options @("Instalar", "Activar", "Desinstalar", "Salir") -Prompt "Seleccione una opcion:"

    switch ($mainOption) {
        "Instalar" { InstallOffice }
        "Activar" { ActivateOffice }
        "Desinstalar" { UninstallOffice }
        "Salir" { exit }
    }

    Write-Host "`nPresione cualquier tecla para volver al menu principal..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}