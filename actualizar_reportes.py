import os
import subprocess

# Ruta de la carpeta compartida en red
carpeta_red = r"\\CONTPAQ\Reportes"

def actualizar_reportes_red():
    print(f"Conectando a la carpeta de red: {carpeta_red}...")
    
    if not os.path.exists(carpeta_red):
        print(f"❌ Error: No se pudo acceder a la ruta {carpeta_red}. Verifica tu conexión de red.")
        return

    # Buscar todos los archivos de Excel (.xlsm, .xlsx) directamente en la carpeta
    archivos_excel = [
        os.path.join(carpeta_red, f) for f in os.listdir(carpeta_red)
        if f.lower().endswith(('.xlsm', '.xlsx')) and not f.startswith('~$')
    ]

    if not archivos_excel:
        print("⚠️ No se encontraron archivos de Excel en la carpeta principal.")
        return

    print(f"Se encontraron {len(archivos_excel)} archivos para actualizar.\n")

    # Script de PowerShell optimizado para abrir, actualizar, guardar y cerrar Excel de forma limpia
    for ruta in archivos_excel:
        nombre_archivo = os.path.basename(ruta)
        print(f"📂 Procesando: {nombre_archivo}...")
        
        # Comando de PowerShell que maneja Excel de forma nativa y robusta
        ps_script = f"""
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        try {{
            $wb = $excel.Workbooks.Open('{ruta}', 3, $false)
            $wb.RefreshAll()
            Start-Sleep -Seconds 12
            $wb.Save()
            $wb.Close($true)
            Write-Output "SUCCESS"
        }} catch {{
            Write-Output "ERROR: $_"
        }} finally {{
            $excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        }}
        """
        
        try:
            resultado = subprocess.run(
                ["powershell", "-Command", ps_script],
                capture_output=True,
                text=True,
                timeout=60
            )
            
            if "SUCCESS" in resultado.stdout:
                print(f"✅ ¡Actualizado y cerrado con éxito: {nombre_archivo}!\n")
            else:
                print(f"⚠️ Error al actualizar {nombre_archivo}: {resultado.stdout.strip()}\n")
                
        except subprocess.TimeoutExpired:
            print(f"⚠️ Tiempo de espera agotado para {nombre_archivo}.\n")
        except Exception as e:
            print(f"⚠️ Error inesperado en {nombre_archivo}: {e}\n")

    print("Proceso finalizado. Todos los reportes de red fueron procesados.")

if __name__ == "__main__":
    actualizar_reportes_red()