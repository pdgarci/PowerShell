################################################################################
# Script para extraer la cotización del dólar y euro blue de 2 sitios web y    #
# salvar los datos en una planilla Excel                                       #
#                                                                              #
# Autor: Pablo D. Garcia                                                       #
# Fecha primera versión: 06/Feb/2023                                           #
# Versión: 1.3                                                                 #
# Fecha: 30/Abr/2023                                                           #
#                                                                              #
# Se agregó la lógica para leer una lista de feriados. Windows task scheduler  #
# no provee la funcionalidad de definir feriados. Se definió una tarea         #
# programada que se ejecuta de lunes a vernes por la noche.                    #
# Ahora además, se lee la lista de feriados nacionales, producida por otro     #
# script, extraídos de la web del gobierno nacional. Si el día de ejecución    #
# automatizada por Windows task scheduler es un feriado, entonces el script    #
# sólo anota en la bitácora eso y no busca cotizaciones ni las graba en la     #
# planilla. Si el archivo de feriados no existe, el script aborta.             #
#                                                                              #
# Versión 1.3.1                                                                #
# Fecha: 09/Oct/2023                                                           #
#                                                                              #
# Se modificó la expresión regular para obtener la cotización del euro         #
#                                                                              #
# Códigos de salida del script:                                                #
# 0 - normal, sin errores                                                      #
# 1 - el archivo con la lista de feriados no existe o no se pudo leer          #
# 2 - El día de ejecución es feriado                                           #
################################################################################

$ti = Get-Date

Function now()
{
    return (Get-Date -Format "yyyy-MM-dd_HH:mm:ss,ffff 'GMT'K")
}

Function now_no_tz()
{
    return (Get-Date -Format "yyyy-MM-dd_HH-mm-ss")
}

Function end([int]$return_code)
{
    $tf = Get-Date
    $delta_t = $tf - $ti
    log "INFO" "Tiempo transcurrido $($delta_t.Hours) horas $($delta_t.Minutes) minutos $($delta_t.Seconds),$($delta_t.Milliseconds) segundos"

    #Remover todas las variables
    Remove-Variable * -ErrorAction SilentlyContinue

    #Ejecutar el garbage collector (para limpiar todos los procesos, particularmente EXCEL
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    #Finalizar la ejecución del script 
    exit $return_code
}

 Function log([string]$nivel,[string]$mensaje)
{
    "$(now) - $($this_computer) - "+$($nivel)+" : "+$($mensaje)|Out-File -append $log_file -encoding ASCII
}

$script_dir        = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$script_name       = $MyInvocation.MyCommand.Name.Split(".",2)[0]   # it is assumed that script file name is "script_name.*"
$this_computer     = $([System.Net.Dns]::GetHostByName((hostname)).HostName)
$main_dir          = $script_dir
$log_file_dir      = "$($main_dir)\$($script_name)\Logs"
$log_file          = "$($log_file_dir)\$(now_no_tz)-$($this_computer)-local.log"
$psv               = $PSVersionTable.PSVersion.ToString()
#$Excel_file        = "F:\Documentos\Pablo\fx\archivos_de_trabajo\Precio_dolar_y_euro_blue-copia_1.xlsx"
$Excel_file        = "F:\Documentos\Pablo\fx\Precio_dolar_y_euro_blue.xlsx"
$holidays_file     = "$($main_dir)\feriados.txt"

# Se crean log_file_dir y log_file

if(! (test-path $log_file_dir))
{
    new-item -path $log_file_dir -itemtype Directory -force -erroraction silentlycontinue > $null
}
	

if (! (test-path $log_file))
{
    new-item -path $log_file -itemtype File -force -erroraction silentlycontinue > $null
}

log "INFO" "Script para obtener cotizaciones dolar y euro blue iniciado"
log "INFO" "Directorio de bitacora '$($log_file_dir)'"
log "INFO" "Archivo de bitacora '$($log_file)'"
log "INFO" "El script se esta ejecutando en la computadora '$($this_computer)'"
log "INFO" "La version de PowerShell es $($psv)"


# Si el módulo PowerHTML no está instalado, lo instala
If (-not (Get-Module -ErrorAction Ignore -ListAvailable PowerHTML)) {
  Write-Verbose "Instalando el modulo PowerHTML para el usuario actual..."
  Install-Module PowerHTML -ErrorAction Stop
}
Import-Module -ErrorAction Stop PowerHTML

# Se verifica la existencia del archivo de entrada con la lista de feriados
if(! (test-path $holidays_file))
{
    log "CRIT" "El archivo '$($holidays_file)' con la lista de feriados no existe"
    end(1)
}

# Se intenta leer el archivo con la lista de feriados
$feriados = ""
$feriados = Get-Content -Path $holidays_file
if ($feriados.Lengh -eq 0)
{
    log "CRIT" "El archivo '$($holidays_file)' con la lista de feriados no se pudo leer"
    end(1)
}

log "INFO" "El archivo '$($holidays_file)' con la lista de feriados se leyo correctamente"

$today = Get-Date -Format "dd/MM/yyyy"
if ($today -in $feriados)
{
    log "WARN" "Hoy '$($today)' es feriado. No se buscara cotizacion alguna."
    end(2)
}

log "INFO" "Hoy '$($today)' no es feriado. Se buscaran las cotizaciones."

$dol_compra = ""
$dol_venta = ""
$status_code = 0
log "INFO" "Contactando 'https://dolarhoy.com/i/cotizaciones/dolar-blue'"
while ($status_code -ne 200)
{
    $response = Invoke-WebRequest -Method Get -Uri https://dolarhoy.com/i/cotizaciones/dolar-blue -SslProtocol Tls13,Tls12 -HttpVersion 3.0
    $status_code = $response.StatusCode
    $is_from_cache = $response.BaseResponse.IsFromCache
}

$html = ConvertFrom-Html -Content $response.Content
$elems = $html.DescendantNodes().Elements('div')

foreach ($elem in $elems)
{
    if ($elem.InnerText -match 'D.lar Blue')
    {
        log "INFO" "Cotizacion obtenida"
        $encontrado = $elem.InnerText -replace 'D.lar Blue'
        $dol_compra = $encontrado.Substring(0,$encontrado.IndexOf('compra',[System.StringComparison]::CurrentCultureIgnoreCase))
        $encontrado = $encontrado -replace $dol_compra
        $encontrado = $encontrado -replace "Compra"
        $dol_venta = $encontrado.Substring(0,$encontrado.IndexOf('Venta',[System.StringComparison]::CurrentCultureIgnoreCase))
        break
    }
}

$dol_venta = [int]$dol_venta
$dol_compra = [int]$dol_compra
log "INFO" "Dolar blue"
log "INFO" ("Compra: $" +$($dol_compra))
log "INFO" ("Venta:  $" +$($dol_venta))

$eur_compra = ""
$eur_venta = ""
$status_code = 0
while ($status_code -ne 200)
{
    log "INFO" "Contactando 'https://tiempofinanciero.com.ar/cotizaciones/euro-blue/'"
    $response = Invoke-WebRequest -Method Get -Uri https://tiempofinanciero.com.ar/cotizaciones/euro-blue/ -SslProtocol Tls13,Tls12 -HttpVersion 3.0
    $status_code = $response.StatusCode
    $is_from_cache = $response.BaseResponse.IsFromCache
}

$html = ConvertFrom-Html -Content $response.Content

$elems = $html.DescendantNodes().Elements('td')

foreach ($elem in $elems)
{
    if ($elem.InnerText -match 'Euro blue')
	{
		log "INFO" "Cotizacion obtenida"
        $encontrado = $elem
        $i = $elems.IndexOf($elem)
        break
	}
}

#$eur_compra = [int]($elems[$i+1].InnerText -replace "\$") # acá reemplazaba '$' por nada
#$eur_venta = [int]($elems[$i+2].InnerText -replace "\$")
$eur_compra = [int]($elems[$i+1].InnerText -replace '\D') # acá reemplaza todos los caracteres que no son dígitos por nada
$eur_venta = [int]($elems[$i+2].InnerText -replace '\D')

log "INFO" "Euro blue"
log "INFO" ("Compra: $"+$($eur_compra))
log "INFO" ("Venta:  $"+$($eur_venta))

# Vuelco de los datos a la planilla Excel

log "INFO" "Abriendo planilla Excel '$($Excel_file)'"

$Excel = New-Object -comobject Excel.Application
$Excel.visible=$false
$Excel_work_book = $Excel.Workbooks.Open($Excel_file)
$Excel_work_sheet = $Excel_work_book.Sheets.Item("Sheet1")

#Columna "A" = "Fecha"
$last_row = $Excel_work_sheet.UsedRange.SpecialCells(11).row
$range = $Excel_work_sheet.Range('A1',"A"+$($last_row))

#Busca si hay una fila en la que ya esté cargada la fecha (hoy) de ejecución del script
$row = 0
$today = ([DateTime]::Today).ToShortDateString()

foreach ($cell in $range.Cells)
{
    if ($cell.Text -eq $today)
    {
        $row = $cell.Row
        break
    }
}

if ($row -gt 0)
{
    log "INFO" "En la planilla Excel se encontro un registro con la fecha de hoy $($today) en la fila $($row)"
    log "INFO" "Se utilizara la celda de esa fila, $($row), para guardar las cotizaciones de hoy $(([DateTime]::Today).ToLongDateString())"
}
else
{
    log "WARN" "En la planilla Excel no se encontro un registro con la fecha de hoy $($today)"

    #Tomo la columna A (Fecha) entera
    $range = $Excel_work_sheet.Range("A1").EntireColumn
    foreach ($cell in $range.Cells)
    {
        #Busca la primer celda en blanco de la columna A (Fecha) y asumo que allí se debe insertar el siguiente
        #registro, que corresponde a la fecha de ejecución del script ($today)
        if ($cell.Text -in ("",$null))
        {
            $row = $cell.Row
            log "INFO" "Se utilizara la primer celda en blanco de la columna A (Fecha) que es la de la fila $($row), para guardar las cotizaciones de hoy $(([DateTime]::Today).ToLongDateString())"
            $Excel_work_sheet.Cells.Item($row,1) = ([DateTime]::Today).ToOADate()
            #$Excel_work_sheet.Cells.Item($row,1) = ([DateTime]::Today).ToString("d/M/yyyy")
            #$Excel_work_sheet.Cells.Item($row,1) = ([DateTime]::Today).ToShortDateString()
            $Excel_work_sheet.Cells.Item($row,1).HorizontalAlignment = -4152
            break
        }
    }
    #exit(1)
}

#Guarda los datos de la cotización del día en el registro correspondiente (fila de fecha de ejecución)
$Excel_work_sheet.Cells.Item($row,2) = $dol_compra
$Excel_work_sheet.Cells.Item($row,3) = $dol_venta
$Excel_work_sheet.Cells.Item($row,4) = $eur_compra
$Excel_work_sheet.Cells.Item($row,5) = $eur_venta
$Excel_work_sheet.Cells.Item($row,2).NumberFormat = "$ #.##0,00"
$Excel_work_sheet.Cells.Item($row,3).NumberFormat = "$ #.##0,00"
$Excel_work_sheet.Cells.Item($row,4).NumberFormat = "$ #.##0,00"
$Excel_work_sheet.Cells.Item($row,5).NumberFormat = "$ #.##0,00"
$Excel_work_sheet.Cells.Item($row,6).formula = "=AVERAGE(D$row,E$row)/AVERAGE(C$row,B$row)"

#Inhabilitación de las alertas de Excel (para que no salga cuadro de diálogo al grabar)
$excel.DisplayAlerts = $false

#Grabado de la planilla Excel
$Excel_work_book.Save()


# Cerrar Excel
$Excel.Quit()

$null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel_work_sheet)
$null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel_work_book)
$null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)

end(0)