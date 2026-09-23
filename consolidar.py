import os
import pandas as pd

# Rutas de las carpetas
carpeta_origen = r"\\CONTPAQ\Reportes"  
carpeta_destino = r"C:\Users\User\Downloads\PYTHON"
archivo_salida = os.path.join(carpeta_destino, "Consolidado_Master.xlsx")

pestanias_a_consolidar = [
    'FacturaCliente', 
    'SolicitudPago', 
    'OrdenCompra', 
    'EdoCuenta', 
    'FacturaCompra', 
    'Gastos'
]

datos_consolidados = {pestaña: [] for pestaña in pestanias_a_consolidar}

print("Iniciando lectura y consolidación de archivos...")

try:
    archivos = [f for f in os.listdir(carpeta_origen) if f.endswith(('.xlsx', '.xlsm'))]
except Exception as e:
    print(f"⚠️ Error al acceder a la carpeta de red {carpeta_origen}: {e}")
    archivos = []

for archivo in archivos:
    ruta_completa = os.path.join(carpeta_origen, archivo)
    nombre_empresa = os.path.splitext(archivo)[0]
    
    print(f"Procesando empresa: {nombre_empresa}...")
    
    try:
        xl = pd.ExcelFile(ruta_completa)
        for pestaña in pestanias_a_consolidar:
            if pestaña in xl.sheet_names:
                df = pd.read_excel(ruta_completa, sheet_name=pestaña)
                if not df.empty:
                    df['EmpresaOrigen'] = nombre_empresa
                    datos_consolidados[pestaña].append(df)
    except Exception as e:
        print(f"⚠️ Error al leer el archivo {archivo}: {e}")

print("\nGuardando archivo maestro consolidado...")

try:
    # Escribir los datos con compatibilidad total para openpyxl moderno
    with pd.ExcelWriter(archivo_salida, engine='openpyxl') as writer:
        for pestaña, lista_dfs in datos_consolidados.items():
            if lista_dfs:
                df_final = pd.concat(lista_dfs, ignore_index=True)
                df_final.to_excel(writer, sheet_name=pestaña, index=False)
                print(f"✔ Pestaña '{pestaña}' guardada con {len(df_final):,} registros totales.")
                
    print(f"\n¡Proceso finalizado con éxito! El archivo maestro se guardó en:\n{archivo_salida}")

except PermissionError:
    print("\n❌ ERROR DE PERMISO: El archivo 'Consolidado_Master.xlsx' está abierto en Excel. Cierra el archivo y vuelve a ejecutar el script.")
except Exception as e:
    print(f"\n❌ Ocurrió un error al guardar: {e}")