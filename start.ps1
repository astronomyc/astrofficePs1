# Funciones

function mainMenu {
    Clear-Host
    Write-Host "-----------------------------------------"
    Write-Host "Select Option:             By: Astronomyc" -ForegroundColor Cyan
    Write-Host "1. Install" -ForegroundColor White
    Write-Host "2. Activate" -ForegroundColor White
    Write-Host "3. Uninstall" -ForegroundColor White
    Write-Host "-----------------------------------------"
}

function productsMenu {
    param (
        [string[]]$Options,
        [bool[]]$SelectedOptions,
        [int]$SelectedLine
    )

    Clear-Host
    Write-Host "Select Products (Use SPACE for select and Enter for confirm)" -ForegroundColor Cyan

    for ($i = 0; $i -lt $Options.Length; $i++) {
        $prefix = if ($SelectedOptions[$i]) { "[X]" } else { "[ ]" }

        if ($i -eq $SelectedLine) {
            Write-Host ("-> " + $prefix + " " + $Options[$i]) -ForegroundColor Green
        }
        else {
            Write-Host ("   " + $prefix + " " + $Options[$i]) -ForegroundColor White
        }
    }
}

function getSelectedOptions {
    param (
        [string[]]$Options,
        [bool[]]$SelectedOptions
    )

    $selected = @()
    for ($i = 0; $i -lt $Options.Length; $i++) {
        if ($SelectedOptions[$i]) {
            $selected += $Options[$i]
        }
    }
    return $selected
}

function GenerateConfig {
    param (
        [string[]]$SelectedProducts,
        [string[]]$AllProducts
    )

    Write-Host "Select Arch"
    Write-Host "1. x64"
    Write-Host "2. x32"
    $Ar = Read-Host "Arch"
    Write-Host "Arch Select" -ForegroundColor Cyan
    switch ($Ar) {
        1 {
            $Arch = 64
            Write-Host "x64" -ForegroundColor Green
        }
        2 {
            $Arch = 32
            Write-Host "x32" -ForegroundColor Green
        }
        default {
            $Arch = 64
            Write-Host "x64" -ForegroundColor Green
        }
    }

    Write-Host "Select Product ID"
    Write-Host "1. MondoVolume 2016"
    Write-Host "2. ProPlusVolume 2016"
    Write-Host "3. StandardVolume 2016"
    Write-Host "4. ProPlus2019Volume"
    Write-Host "5. Standard2019Volume"
    Write-Host "6. ProPlus2021Volume"
    Write-Host "7. Standard2021Volume"
    Write-Host "8. ProPlus2024Volume"
    Write-Host "9. Standard2024Volume"
    Write-Host "10. O365ProPlusRetail"
    $ProductIdOption = Read-Host "Product ID"
    Write-Host "Product Select" -ForegroundColor Cyan
    switch ($ProductIdOption) {
        1 {
            $ProductId = "MondoVolume"
            Write-Host "MondoVolume 2016 Selected" -ForegroundColor Green
        }
        2 {
            $ProductId = "ProPlusVolume"
            Write-Host "ProPlusVolume 2016 Selected" -ForegroundColor Green
        }
        3 {
            $ProductId = "StandardVolume"
            Write-Host "StandardVolume 2016 Selected" -ForegroundColor Green
        }
        4 {
            $ProductId = "ProPlus2019Volume"
            Write-Host "ProPlus2019Volume Selected" -ForegroundColor Green
        }
        5 {
            $ProductId = "Standard2019Volume"
            Write-Host "Standard2019Volume Selected" -ForegroundColor Green
        }
        6 {
            $ProductId = "ProPlus2021Volume"
            Write-Host "ProPlus2021Volume Selected" -ForegroundColor Green
        }
        7 {
            $ProductId = "Standard2021Volume"
            Write-Host "Standard2021Volume Selected" -ForegroundColor Green
        }
        8 {
            $ProductId = "ProPlus2024Volume"
            Write-Host "ProPlus2024Volume Selected" -ForegroundColor Green
        }
        9 {
            $ProductId = "Standard2024Volume"
            Write-Host "Standard2024Volume Selected" -ForegroundColor Green
        }
        10 {
            $ProductId = "O365ProPlusRetail"
            Write-Host "O365ProPlusRetail Selected" -ForegroundColor Green
        }
        default {
            $ProductId = "ProPlus2021Volume"
            Write-Host "Invalid selection. Defaulting to ProPlus2021Volume" -ForegroundColor Green
        }
    }

    Write-Host "Select Language"
    Write-Host "1. es-es"
    Write-Host "2. en-us"
    $LanguageOption = Read-Host "Language"
    Write-Host "Language Select" -ForegroundColor Cyan
    switch ($LanguageOption) {
        1 {
            $Language = "es-es"
            Write-Host "es-es Selected" -ForegroundColor Green
        }
        2 {
            $Language = "en-us"
            Write-Host "en-us Selected" -ForegroundColor Green
        }
        default {
            $Language = "es-es"
            Write-Host "es-es Selected" -ForegroundColor Green
        }
    }
    $excludedApps = $AllProducts | Where-Object { $_ -notin $SelectedProducts }
    $nonProjectVisioSelected = $SelectedProducts | Where-Object { $_ -ne "Project" -and $_ -ne "Visio" }
    $xmlContent = @"
<Configuration>`n
    <Add OfficeClientEdition="$Arch" Channel="PerpetualVL2021">`n
"@

    # Verificar si hay productos seleccionados excluyendo Project y Visio
    if ($nonProjectVisioSelected.Count -gt 0) {
        $xmlContent += @"
        <Product ID="$ProductId">`n
            <Language ID="$Language" />`n
"@
        foreach ($app in $excludedApps) {
            if ($app -eq "Project" -or $app -eq "Visio") {
                continue
            }
            $xmlContent += "            <ExcludeApp ID=""$app"" />`n"
        }
        $xmlContent += @"
        </Product>
"@
    }

    # Añadir la seccion de Project si esta seleccionado
    if ($SelectedProducts -contains "Project") {
        $xmlContent += @"
        <Product ID="ProjectProRetail">`n
            <Language ID="$Language" />`n
        </Product>`n
"@
    }

    # Añadir la seccion de Visio si esta seleccionado
    if ($SelectedProducts -contains "Visio") {
        $xmlContent += @"
        <Product ID="VisioProRetail">`n
            <Language ID="$Language" />`n
        </Product>`n
"@
    }

    $xmlContent += @"
    </Add>

    <Updates Enabled="False"  />

</Configuration>
"@
    $xmlContent | Out-File -FilePath "config.xml" -Encoding utf8
}

function Install {
    $officeProducts = @("Word", "Excel", "PowerPoint", "Access", "Publisher", "OneNote", "Outlook", "Teams", "OneDrive", "Lync", "Project", "Visio")
    $selectedOptions = @(,$false) * $officeProducts.Length
    $SelectedLine = 0

    while ($true) {
        productsMenu -Options $officeProducts -SelectedOptions $selectedOptions -SelectedLine $SelectedLine
        $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown").VirtualKeyCode

        if ($key -eq 38) {
            $SelectedLine = [Math]::Max(0, $SelectedLine - 1)
        }
        elseif ($key -eq 40) {
            $SelectedLine = [Math]::Min($officeProducts.Length - 1, $SelectedLine + 1)
        }
        elseif ($key -eq 32) {
            $selectedOptions[$SelectedLine] = -not $selectedOptions[$SelectedLine]
        }
        elseif ($key -eq 13) {
            $selectedProducts = getSelectedOptions -Options $officeProducts -SelectedOptions $selectedOptions
            Clear-Host
            Write-Host "Select Products:" -ForegroundColor Cyan
            foreach ($product in $selectedProducts) {
                Write-Host $product -ForegroundColor Green
            }
            GenerateConfig -SelectedProducts $selectedProducts -AllProducts $officeProducts
            (New-Object System.Net.WebClient).DownloadFile('TODO', "$env:TEMP\TODO")
            & "$env:TEMP\TODO"
            TODO
            break
        }
    }
}

function selectMainMenu {
    param (
        [int]$Option
    )

    switch ($Option) {
        1 {
            Install
        }
        2 {
            Activate
        }
        3 {
            Write-Host "Coming Soon"
        }
        default {
            mainMenu
            $Option = Read-Host "Option"
            selectMainMenu -Option $Option
        }
    }
}

function Activate {
    Invoke-RestMethod https://get.activated.win | Invoke-Expression
}

# Inciar
mainMenu
$Option = Read-Host "Option"
selectMainMenu -Option $Option
