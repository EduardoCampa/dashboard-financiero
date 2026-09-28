import os
import subprocess
import time

# Ruta de la carpeta compartida en red
carpeta_red = r"\\CONTPAQ\Reportes\ARCHIVOS PROG"


def liberar_procesos_excel_bloqueados():
    """Cierra cualquier proceso colgado de Excel en Windows."""
    try:
        subprocess.run(
            ["taskkill", "/f", "/im", "excel.exe", "/t"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except Exception:
        pass


def es_archivo_bloqueado(ruta):
    """Detecta si el archivo está en uso por otro usuario o si existe su archivo de propietario ~$."""
    dir_name, file_name = os.path.split(ruta)
    archivo_propietario = os.path.join(dir_name, f"~${file_name}")

    # Si existe el archivo oculto ~$CIVLA.xlsm, está abierto en la red
    if os.path.exists(archivo_propietario):
        return True

    # Intentar abrir el archivo de forma exclusiva en escritura
    try:
        with open(ruta, "a+b") as f:
            pass
        return False
    except IOError:
        return True


def actualizar_reportes_red():
    print("🧹 Liberando procesos de Excel colgados en segundo plano...")
    liberar_procesos_excel_bloqueados()

    print(f"Conectando a la carpeta de red: {carpeta_red}...")

    if not os.path.exists(carpeta_red):
        print(
            f"❌ Error: No se pudo acceder a la ruta {carpeta_red}. Verifica tu conexión de red."
        )
        return

    # Buscar archivos de Excel (.xlsm, .xlsx) ignorando los temporales ~$
    archivos_excel = [
        os.path.join(carpeta_red, f)
        for f in os.listdir(carpeta_red)
        if f.lower().endswith((".xlsm", ".xlsx")) and not f.startswith("~$")
    ]

    if not archivos_excel:
        print("⚠️ No se encontraron archivos de Excel en la carpeta principal.")
        return

    print(f"Se encontraron {len(archivos_excel)} archivos para actualizar.\n")

    for ruta in archivos_excel:
        nombre_archivo = os.path.basename(ruta)
        print(f"📂 Procesando: {nombre_archivo}...")

        # 1. Verificar si el archivo está bloqueado antes de llamar a Excel
        if es_archivo_bloqueado(ruta):
            print(
                f"⚠️ El archivo {nombre_archivo} está abierto por otro usuario en la red. Omitiendo actualización para evitar congelamiento.\n"
            )
            continue

        # 2. Script de PowerShell con supresión total de diálogos emergentes y macros de seguridad
        ps_script = f"""
        $ErrorActionPreference = 'Stop'
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.AskToUpdateLinks = $false
        $excel.AlertBeforeOverwriting = $false
        $excel.AutomationSecurity = 3 # Desactivar avisos de seguridad de macros flotantes

        try {{
            # Abrir archivo forzando ignorar alertas de enlaces y bloqueos
            $wb = $excel.Workbooks.Open('{ruta}', 0, $false, 5, '', '', $true)
            
            if ($wb.ReadOnly) {{
                Write-Output "LOCKED"
                $wb.Close($false)
            }} else {{
                $wb.RefreshAll()
                Start-Sleep -Seconds 8
                $wb.Save()
                $wb.Close($true)
                Write-Output "SUCCESS"
            }}
        }} catch {{
            Write-Output "ERROR: $_"
        }} finally {{
            $excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }}
        """

        try:
            resultado = subprocess.run(
                ["powershell", "-ExecutionPolicy", "Bypass", "-Command", ps_script],
                capture_output=True,
                text=True,
                timeout=35,  # Reducimos el tiempo máximo a 35 segundos para que no se atore
            )

            out_text = resultado.stdout.strip()

            if "SUCCESS" in out_text:
                print(f"✅ ¡Actualizado y cerrado con éxito: {nombre_archivo}!\n")
            elif "LOCKED" in out_text:
                print(
                    f"⚠️ {nombre_archivo} está en modo Solo Lectura en la red. Omitido.\n"
                )
            else:
                print(
                    f"⚠️ Alerta al procesar {nombre_archivo}: {out_text}\n"
                )

        except subprocess.TimeoutExpired:
            print(
                f"⚠️ Tiempo agotado para {nombre_archivo}. Se forzó la liberación para no detener los demás reportes.\n"
            )
            liberar_procesos_excel_bloqueados()
            time.sleep(2)
        except Exception as e:
            print(f"⚠️ Error inesperado en {nombre_archivo}: {e}\n")

    liberar_procesos_excel_bloqueados()
    print("✨ Proceso finalizado. Todos los reportes de red fueron procesados.")


if __name__ == "__main__":
    actualizar_reportes_red()