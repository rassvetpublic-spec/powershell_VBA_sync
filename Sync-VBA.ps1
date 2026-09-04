# -*- coding: utf-8 -*-
<#
    Sync-VBA.ps1  v2025.11.17r3
    Скрипт синхронизации VBA-модулей Excel:
    - Экспорт модулей/классов/форм в папку VBA (UTF-8 BOM)
    - Импорт модулей/классов/форм из папки VBA
    - Поддержка x86 Excel через перезапуск 32-битного PowerShell
#>

param(
    [int]$Mode = 0,
    [string]$ProjectPath = (Get-Location)
)

# Глобальный путь к лог-файлу
$script:SyncVba_LogFile = $null
 
<# ======================= ЛОГИРОВАНИЕ ======================= #>
function Write-Log {
    <#
        Многострочный комментарий:
        Функция логирования в консоль и файл SyncVBA.log.
        Цветной вывод в консоль, в файл пишем всегда текст с меткой времени.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ConsoleColor]$Color = [ConsoleColor]::Gray
    )

    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $line = "[{0}] {1}" -f $timestamp, $Message

    Write-Host $line -ForegroundColor $Color

    if ($script:SyncVba_LogFile) {
        Add-Content -Encoding UTF8 -Path $script:SyncVba_LogFile -Value $line
    }
}

<# ======================= КОДИРОВКИ / MOJIBAKE ======================= #>
function Test-Mojibake {
    <#
        Проверка строки на типичные "кракозябры" после перепутанной
        UTF-8 / ANSI кодировки.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text
    )

    if ([string]::IsNullOrEmpty($Text)) {
        return $false
    }

    return ($Text -match '[ÃÐÑâ€“â€”â€œâ€â€˜â€™¢™€]')
}

function Fix-Mojibake {
    <#
        Попытка "починить" кракозябры:
        считаем, что текст ошибочно прочитали как 1252 вместо UTF-8,
        и перекодируем обратно.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text
    )

    $bytes = [System.Text.Encoding]::GetEncoding(1252).GetBytes($Text)
    return [System.Text.Encoding]::UTF8.GetString($bytes)
}

function Write-UTF8BOM {
    <#
        Запись текста в файл с явным UTF-8 BOM,
        как требуется для проекта.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Text
    )

    $encoding = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText($Path, $Text, $encoding)
}

function Convert-TextFile-ToUtf8Bom {
    <#
        Чтение текстового файла в системной ANSI-кодировке,
        починка кракозябр (если есть) и запись в UTF-8 BOM.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -Path $Path)) {
        return
    }

    $ansiEncoding = [System.Text.Encoding]::Default
    $raw = [System.IO.File]::ReadAllText($Path, $ansiEncoding)
    if (Test-Mojibake -Text $raw) {
        $raw = Fix-Mojibake -Text $raw
    }

    Write-UTF8BOM -Path $Path -Text $raw
}

<# ======================= СРЕДА / x86 ПЕРЕЗАПУСК ======================= #>
function Write-EnvironmentInfo {
    <#
        Логируем архитектуру PowerShell и Excel.
        Используем ключ реестра Excel 2016 (16.0) для примерной оценки.
    #>
    [CmdletBinding()]
    param()

    $psArch = if ([Environment]::Is64BitProcess) { "x64" } else { "x86" }
    $excelArch = ""

    try {
        $key = "HKLM:\SOFTWARE\Microsoft\Office\16.0\Excel\InstallRoot"
        if (Test-Path -Path $key) {
            $path = (Get-ItemProperty -Path $key).Path
            $excelArch = if ($path -match "Program Files \(x86\)") { "x86" } else { "x64" }
        }
    }
    catch {
        $excelArch = ""
    }

    Write-Log ("📊 Среда: Excel={0}, PowerShell={1}" -f $excelArch, $psArch)
}

function Stop-ExcelAll {
    <#
        Принудительное завершение всех процессов Excel.
        Используется в режиме KillExcel.
    #>
    [CmdletBinding()]
    param()

    Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue |
        Stop-Process -Force -ErrorAction SilentlyContinue
}

function Invoke-32BitSelf {
    <#
        Перезапуск текущего скрипта в 32-битной версии PowerShell
        для работы с 32-битным Excel.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [int]$Mode,

        [Parameter(Mandatory = $true)]
        [string]$ProjectPath
    )

    if (-not [Environment]::Is64BitProcess) {
        return
    }

    $wowPath = Join-Path -Path $env:SystemRoot -ChildPath 'SysWOW64\WindowsPowerShell\v1.0\powershell.exe'
    if (-not (Test-Path -Path $wowPath)) {
        return
    }

    Write-Log "Перезапуск в 32-битной версии PowerShell для совместимости с Excel..." ([ConsoleColor]::Yellow)

    $argumentList = @(
        '-NoExit',                      # для отладки оставляем окно открытым
        '-ExecutionPolicy','Bypass',
        '-NoProfile',
        '-File', "`"$PSCommandPath`"",
        '-Mode', $Mode,
        '-ProjectPath', "`"$ProjectPath`""
    )

    Start-Process -FilePath $wowPath -ArgumentList $argumentList -Wait
    exit
}

<# ======================= ВЫБОР КНИГИ EXCEL ======================= #>
function Select-Workbook {
    <#
        Логика выбора Excel-книги:
        1) если книги уже открыты — даём выбрать;
        2) если нет — ищем .xlsm в ProjectPath;
        3) если не нашли — просим путь руками.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Excel,

        [Parameter(Mandatory = $true)]
        [string]$ProjectPath
    )

    try {
        $workbooks = $Excel.Workbooks
        $wbCount   = $workbooks.Count
    }
    catch {
        Write-Log ("⚠ Не удалось получить коллекцию Workbooks: {0}" -f $_.Exception.Message) ([ConsoleColor]::Red)
        return $null
    }

    if ($wbCount -eq 0) {
        # Нет открытых книг — ищем .xlsm в каталоге проекта
        $xlsmFiles = Get-ChildItem -Path $ProjectPath -Filter '*.xlsm' -ErrorAction SilentlyContinue
        if ($xlsmFiles.Count -eq 1) {
            return $Excel.Workbooks.Open($xlsmFiles.FullName)
        }
        elseif ($xlsmFiles.Count -gt 1) {
            Write-Host "`nНайдено несколько файлов Excel:" -ForegroundColor Yellow
            for ($index = 0; $index -lt $xlsmFiles.Count; $index++) {
                Write-Host ("  {0}. {1}" -f ($index + 1), $xlsmFiles[$index].Name)
            }
            $selection = Read-Host "Введите номер файла"
            $selectedIndex = [int]$selection - 1
            if ($selectedIndex -ge 0 -and $selectedIndex -lt $xlsmFiles.Count) {
                return $Excel.Workbooks.Open($xlsmFiles[$selectedIndex].FullName)
            }
            else {
                Write-Log "❌ Неверный номер файла при выборе .xlsm" ([ConsoleColor]::Red)
                return $null
            }
        }
        else {
            Write-Host "Нет открытых книг и .xlsm не найдено. Укажи путь:" -ForegroundColor Yellow
            $path = Read-Host "Полный путь к .xlsm"
            if ([string]::IsNullOrWhiteSpace($path)) {
                Write-Log "❌ Путь к книге не указан." ([ConsoleColor]::Red)
                return $null
            }
            return $Excel.Workbooks.Open($path)
        }
    }
    elseif ($wbCount -eq 1) {
        # Одна открытая книга
        return $workbooks.Item(1)
    }
    else {
        # Несколько открытых книг — даём выбрать
        Write-Host "`nНайдено несколько открытых книг:" -ForegroundColor Yellow
        for ($index = 1; $index -le $wbCount; $index++) {
            $wb = $workbooks.Item($index)
            Write-Host ("  {0}. {1}" -f $index, $wb.Name)
        }
        $selection = Read-Host "Введите номер файла"
        $selectedIndex = [int]$selection
        if ($selectedIndex -ge 1 -and $selectedIndex -le $wbCount) {
            return $workbooks.Item($selectedIndex)
        }
        else {
            Write-Log "❌ Неверный номер файла при выборе открытой книги" ([ConsoleColor]::Red)
            return $null
        }
    }
}

<# ======================= СЕССИЯ EXCEL ======================= #>
function Start-ExcelSession {
    <#
        Создаём или находим Excel, выбираем книгу,
        подготавливаем папку VBA и возвращаем объект с параметрами сессии.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProjectPath
    )

    Write-Log "🧭 Поиск активного Excel..." ([ConsoleColor]::Gray)

    $excel = $null
    $createdNewExcel = $false

    try {
        $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
        Write-Log "📎 Подключились к активному Excel." ([ConsoleColor]::Green)
    }
    catch {
        $running = Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue
        if ($running) {
            Write-Log "⚠ Excel запущен, но COM недоступен — создаём новый экземпляр." ([ConsoleColor]::Yellow)
        }
        else {
            Write-Log "⚠ Excel не найден — создаём новый экземпляр." ([ConsoleColor]::Yellow)
        }
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true
        $createdNewExcel = $true
    }

    $excel.DisplayAlerts  = $false
    $excel.EnableEvents   = $false
    $excel.ScreenUpdating = $false
    $excel.Interactive    = $false

    $workbook = Select-Workbook -Excel $excel -ProjectPath $ProjectPath

    if (-not $workbook) {
        Write-Log "❌ Не удалось выбрать книгу Excel." ([ConsoleColor]::Red)
        return $null
    }

    $workbookName     = $workbook.Name
    $workbookBaseName = [System.IO.Path]::GetFileNameWithoutExtension($workbookName)

    Write-Log ("📘 Активная книга: {0}" -f $workbookName)

    $exportPath = Join-Path -Path $ProjectPath -ChildPath 'VBA'
    if (-not (Test-Path -Path $exportPath)) {
        New-Item -ItemType Directory -Path $exportPath | Out-Null
    }

    return [pscustomobject]@{
        Excel             = $excel
        Workbook          = $workbook
        WorkbookName      = $workbookName
        WorkbookBaseName  = $workbookBaseName
        ProjectPath       = $ProjectPath
        ExportPath        = $exportPath
        CreatedNewExcel   = $createdNewExcel
    }
}

function Stop-ExcelSession {
    <#
        Аккуратно сохраняем книгу, восстанавливаем параметры Excel,
        и закрываем Excel только если мы его сами создавали.

        Логика:
        - если Excel был уже запущен (CreatedNewExcel = $false):
            * книгу сохраняем, но НЕ закрываем;
            * Excel не закрываем;
        - если Excel создал скрипт (CreatedNewExcel = $true):
            * сохраняем книгу;
            * закрываем книгу;
            * делаем Quit() Excel.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Session
    )

    if (-not $Session) {
        return
    }

    $excel           = $Session.Excel
    $workbook        = $Session.Workbook
    $createdNewExcel = $Session.CreatedNewExcel

    try {
        if ($workbook -ne $null) {
            try {
                Write-Log "💾 Сохраняем книгу..." ([ConsoleColor]::Gray)
                $workbook.Save()
            }
            catch {
                Write-Log ("⚠ Ошибка при сохранении книги: {0}" -f $_.Exception.Message) ([ConsoleColor]::Red)
            }

            if ($createdNewExcel) {
                # Книга и Excel были созданы скриптом — закрываем книгу
                try {
                    $workbook.Close($true) | Out-Null
                    Write-Log "📕 Книга закрыта (скрипт сам её открывал)." ([ConsoleColor]::DarkGray)
                }
                catch {
                    Write-Log ("⚠ Ошибка при закрытии книги: {0}" -f $_.Exception.Message) ([ConsoleColor]::Red)
                }
            }
            else {
                # Книга была открыта до запуска скрипта — оставляем открытой
                Write-Log "🔁 Книга была открыта пользователем — оставляем её открытой." ([ConsoleColor]::DarkGray)
            }
        }
    }
    finally {
        if ($excel -ne $null) {
            try {
                $excel.DisplayAlerts  = $true
                $excel.EnableEvents   = $true
                $excel.ScreenUpdating = $true
                $excel.Interactive    = $true

                if ($createdNewExcel) {
                    $excel.Quit()
                    Write-Log "✅ Закрыт экземпляр Excel, созданный скриптом." ([ConsoleColor]::DarkGray)
                }
                else {
                    Write-Log "🔁 Excel был запущен ранее — оставляем его работать." ([ConsoleColor]::DarkGray)
                }
            }
            catch {
                Write-Log ("⚠ Ошибка при финализации Excel: {0}" -f $_.Exception.Message) ([ConsoleColor]::Red)
            }

            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }
    }
}

<# ======================= ЭКСПОРТ / ИМПОРТ ======================= #>
function Export-VBAModules {
    <#
        Экспорт всех модулей, классов и форм VBA
        в папку VBA с сохранением UTF-8 BOM.
        Имена файлов: <ИмяКнигиБезРасширения>_<ИмяМодуля>.bas/.cls/.frm
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Session
    )

    $workbook          = $Session.Workbook
    $exportPath        = $Session.ExportPath
    $workbookBaseName  = $Session.WorkbookBaseName

    Write-Log ">>> Экспорт VBA-компонентов..." ([ConsoleColor]::Gray)

    try {
        $vbComponents = @($workbook.VBProject.VBComponents | Where-Object { $_.Type -ne 100 })
    }
    catch {
        Write-Log ("❌ Ошибка доступа к VBProject (проверь 'Trust access to the VBA project'): {0}" -f $_.Exception.Message) ([ConsoleColor]::Red)
        return
    }

    $total = $vbComponents.Count
    $index = 0

    foreach ($vbComponent in $vbComponents) {
        $index++
        $percent = if ($total -gt 0) { [int](($index / $total) * 100) } else { 0 }

        Write-Progress -Activity "Экспорт VBA" -Status $vbComponent.Name -PercentComplete $percent

        try {
            switch ($vbComponent.Type) {
                1 { $extension = ".bas" }  # стандартный модуль
                2 { $extension = ".cls" }  # класс
                3 { $extension = ".frm" }  # форма
                default { $extension = ".bas" }
            }

            $fileName   = "{0}_{1}{2}" -f $workbookBaseName, $vbComponent.Name, $extension
            $targetPath = Join-Path -Path $exportPath -ChildPath $fileName

            if ($extension -in @(".bas", ".cls")) {
                $lineCount = $vbComponent.CodeModule.CountOfLines
                if ($lineCount -gt 0) {
                    $codeText = $vbComponent.CodeModule.Lines(1, $lineCount)
                    if (Test-Mojibake -Text $codeText) {
                        $codeText = Fix-Mojibake -Text $codeText
                    }
                    Write-UTF8BOM -Path $targetPath -Text $codeText
                    Write-Log ("✔ Экспортирован модуль: {0}" -f $fileName) ([ConsoleColor]::Green)
                }
            }
            elseif ($extension -eq ".frm") {
                $vbComponent.Export($targetPath)
                Convert-TextFile-ToUtf8Bom -Path $targetPath
                Write-Log ("✔ Экспортирована форма: {0}" -f $fileName) ([ConsoleColor]::Green)
            }
        }
        catch {
            Write-Log ("⚠ Ошибка при экспорте {0}: {1}" -f $vbComponent.Name, $_.Exception.Message) ([ConsoleColor]::Red)
        }
    }

    Write-Progress -Activity "Экспорт VBA" -Completed -Status "Готово"
    Write-Log "✅ Все модули успешно экспортированы." ([ConsoleColor]::Cyan)
}

function Import-VBAModules {
    <#
        Импорт модулей, классов и форм VBA из папки VBA в книгу.
        Ожидаемый шаблон имени файла:
        <ИмяКнигиБезРасширения>_<ИмяМодуля>.bas/.cls/.frm
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Session
    )

    $workbook          = $Session.Workbook
    $exportPath        = $Session.ExportPath
    $workbookBaseName  = $Session.WorkbookBaseName

    Write-Log ">>> Импорт VBA-компонентов..." ([ConsoleColor]::Gray)

    # Вместо -Include используем -Filter + Where-Object по расширению
    $files = Get-ChildItem -Path $exportPath -File -Filter ("{0}_*" -f $workbookBaseName) -ErrorAction SilentlyContinue |
             Where-Object { $_.Extension.ToLowerInvariant() -in '.bas', '.cls', '.frm' }

    Write-Log ("🔍 Найдено файлов для импорта: {0}" -f ($files.Count)) ([ConsoleColor]::DarkGray)

    if (-not $files -or $files.Count -eq 0) {
        Write-Log "⚠ Подходящих файлов не найдено (проверь папку VBA и имена файлов вида <Книга>_<Модуль>.bas)." ([ConsoleColor]::Yellow)
        return
    }

    foreach ($file in $files) {
        $fileBaseName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
        $extension    = $file.Extension.ToLowerInvariant()
        $prefix       = "{0}_" -f $workbookBaseName

        if (-not $fileBaseName.StartsWith($prefix)) {
            # На всякий случай, защитный фильтр
            continue
        }

        # Имя модуля = всё после "<ИмяКниги>_"
        $moduleName = $fileBaseName.Substring($prefix.Length)
        if ([string]::IsNullOrWhiteSpace($moduleName)) {
            continue
        }

        try {
            if ($extension -in @(".bas", ".cls")) {
                # Читаем текст модуля из файла
                $text = Get-Content -Raw -Encoding UTF8 -Path $file.FullName
                if (Test-Mojibake -Text $text) {
                    $text = Fix-Mojibake -Text $text
                }

                # Ищем существующий компонент с таким именем
                $vbComponent = $workbook.VBProject.VBComponents | Where-Object { $_.Name -eq $moduleName }
                if (-not $vbComponent) {
                    $componentType = if ($extension -eq ".cls") { 2 } else { 1 }  # 1=standard, 2=class
                    $vbComponent = $workbook.VBProject.VBComponents.Add($componentType)
                    $vbComponent.Name = $moduleName
                }

                $codeModule = $vbComponent.CodeModule
                $linesCount = $codeModule.CountOfLines
                if ($linesCount -gt 0) {
                    $codeModule.DeleteLines(1, $linesCount)
                }

                $codeModule.AddFromString($text)

                Write-Log ("✔ Импортирован модуль: {0} ({1})" -f $moduleName, $file.Name) ([ConsoleColor]::Green)
            }
            elseif ($extension -eq ".frm") {
                # Формы: удаляем старую, импортируем новую
                $existing = $workbook.VBProject.VBComponents | Where-Object { $_.Name -eq $moduleName }
                if ($existing) {
                    $workbook.VBProject.VBComponents.Remove($existing)
                }

                $null = $workbook.VBProject.VBComponents.Import($file.FullName)

                # Поддержка .frx рядом с .frm
                $frxPath = [System.IO.Path]::ChangeExtension($file.FullName, ".frx")
                if (Test-Path -Path $frxPath) {
                    $targetFrx = Join-Path -Path ([System.IO.Path]::GetDirectoryName($workbook.FullName)) -ChildPath ([System.IO.Path]::GetFileName($frxPath))
                    Copy-Item -Path $frxPath -Destination $targetFrx -Force
                }

                Write-Log ("✔ Импортирована форма: {0} ({1})" -f $moduleName, $file.Name) ([ConsoleColor]::Green)
            }
        }
        catch {
            Write-Log ("⚠ Ошибка при импорте {0} из {1}: {2}" -f $moduleName, $file.Name, $_.Exception.Message) ([ConsoleColor]::Red)
        }
    }

    Write-Log "✅ Импорт завершён, книга сохранена." ([ConsoleColor]::Cyan)
}

<# ======================= ОТКРЫТИЕ РЕЗУЛЬТАТОВ ЭКСПОРТА ======================= #>
function Open-VbaInEditor {
    <#
        Открываем результаты экспорта:
        1) Если установлен Notepad++ по пути C:\Program Files\Notepad++\notepad++.exe,
           открываем в нём все .bas-файлы текущей книги.
        2) Если Notepad++ не найден или .bas нет,
           просто открываем папку VBA в Проводнике.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExportPath,

        [Parameter(Mandatory = $true)]
        [string]$WorkbookBaseName
    )

    $notepadPath = "C:\Program Files\Notepad++\notepad++.exe"

    $patternBas = "{0}_*.bas" -f $WorkbookBaseName
    $basFiles = Get-ChildItem -Path $ExportPath -Filter $patternBas -File -ErrorAction SilentlyContinue

    if ((Test-Path -Path $notepadPath) -and $basFiles -and $basFiles.Count -gt 0) {
        $args = $basFiles.FullName
        Start-Process -FilePath $notepadPath -ArgumentList $args
        Write-Log ("📄 Открыто в Notepad++ файлов: {0}" -f $basFiles.Count) ([ConsoleColor]::DarkGray)
    }
    else {
        Start-Process -FilePath 'explorer.exe' -ArgumentList "`"$ExportPath`""
        Write-Log "📂 Открыт каталог VBA в Проводнике." ([ConsoleColor]::DarkGray)
    }
}

<# ======================= ФИНАЛЬНОЕ СООБЩЕНИЕ ======================= #>
function Show-FinishMessage {
    <#
        Финальное сообщение и пауза, чтобы окно не закрывалось мгновенно.
    #>
    [CmdletBinding()]
    param()

    Write-Host "`n=== Работа завершена. Нажми любую клавишу для выхода... ===" -ForegroundColor Gray
    Pause
}

<# ======================= ГЛАВНАЯ ФУНКЦИЯ ======================= #>
function Invoke-SyncVbaMain {
    <#
        Главная точка входа: меню, перезапуск в x86, запуск экспорта/импорта.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [int]$Mode,

        [Parameter(Mandatory = $true)]
        [string]$ProjectPath
    )

    $script:SyncVba_LogFile = Join-Path -Path $ProjectPath -ChildPath 'SyncVBA.log'
    Add-Content -Encoding UTF8 -Path $script:SyncVba_LogFile -Value "`n=== Run $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ==="

    chcp 65001 > $null
    [Console]::InputEncoding  = [System.Text.Encoding]::UTF8
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8

    $today = Get-Date
    if ($today.Month -eq 11 -and $today.Day -eq 11) {
        Write-Log "🎂 С днём рождения, инженер Александр!" ([ConsoleColor]::Magenta)
        [Console]::Beep(880,150); [Console]::Beep(988,150); [Console]::Beep(1047,250)
    }

    Write-EnvironmentInfo

    $effectiveMode = $Mode
    if ($effectiveMode -eq 0) {
        Write-Host "`n 1-Экспорт  2-Импорт  3-Оба  4-KillExcel" -ForegroundColor Cyan
        $inputValue = Read-Host "Введите режим"
        [void][int]::TryParse($inputValue, [ref]$effectiveMode)
    }

    switch ($effectiveMode) {
        1 { Write-Log "🚀 Режим: ЭКСПОРТ" ([ConsoleColor]::Cyan) }
        2 { Write-Log "🚀 Режим: ИМПОРТ" ([ConsoleColor]::Cyan) }
        3 { Write-Log "🚀 Режим: ЭКСПОРТ+ИМПОРТ" ([ConsoleColor]::Cyan) }
        4 {
            Write-Log "💀 Завершаем все процессы Excel..." ([ConsoleColor]::Yellow)
            Stop-ExcelAll
            Write-Log "✅ Все экземпляры Excel завершены." ([ConsoleColor]::Green)
            Show-FinishMessage
            return
        }
        Default {
            Write-Log "❌ Неизвестный режим." ([ConsoleColor]::Red)
            Show-FinishMessage
            return
        }
    }

    Invoke-32BitSelf -Mode $effectiveMode -ProjectPath $ProjectPath

    $session = Start-ExcelSession -ProjectPath $ProjectPath
    if (-not $session) {
        Show-FinishMessage
        return
    }

    $exportDone = $false
    $importDone = $false

    try {
        if ($effectiveMode -eq 1 -or $effectiveMode -eq 3) {
            Export-VBAModules -Session $session
            $exportDone = $true
        }

        if ($effectiveMode -eq 2 -or $effectiveMode -eq 3) {
            Import-VBAModules -Session $session
            $importDone = $true
        }
    }
    finally {
        Stop-ExcelSession -Session $session
    }

    if ($exportDone -and -not $importDone) {
        Open-VbaInEditor -ExportPath $session.ExportPath -WorkbookBaseName $session.WorkbookBaseName
    }

    Show-FinishMessage
}

# Автоматический запуск, если скрипт запускают как .\Sync-VBA.ps1,
# а не dot-source (в обёртках Export-VBA.ps1 / Import-VBA.ps1).
if ($MyInvocation.InvocationName -ne '.') {
    Invoke-SyncVbaMain -Mode $Mode -ProjectPath $ProjectPath
}
