# Generador de Building Blocks - V6

## Cambios incluidos

- Titulo de la app: **📦 Generador de Building Blocks**.
- Nombre de pestaña del navegador: **Generador de Building Blocks**.
- Opcion para descargar una **plantilla Excel en blanco** compatible con el input esperado.
- Se eliminaron del panel izquierdo los campos Gerente de marca, Evidencia / soporte y Aprobador.
- Orden de filtros de alcance: Socio Comercial, Marca BU, Division, Fuerza de Venta, Territorio.
- En informacion clave del SKU se muestra congelamiento solo en meses.
- La ZEF se calcula estrictamente con el congelamiento en meses:
  - 1 mes de congelamiento: solo el primer mes abierto queda en ZEF.
  - 2 meses de congelamiento: los dos primeros meses abiertos quedan en ZEF.
  - El mes siguiente a ZEF queda en ZC.
  - Los meses posteriores quedan en ZF.

## Ejecutar

1. Descomprimir el ZIP.
2. Abrir la carpeta `BB_Streamlit_App_V6_Building_Blocks`.
3. Dar doble clic en `EJECUTAR_APP_V6.bat`.

O ejecutar por PowerShell:

```powershell
cd "C:\Users\Juan Carlos Sanchez\Downloads\BB_Streamlit_App_V6_Building_Blocks"
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
.\EJECUTAR_APP_V6.ps1
```

## Archivos esperados

- `app.py`
- `bb_core.py`
- `requirements.txt`
- `EJECUTAR_APP_V6.ps1`
- `EJECUTAR_APP_V6.bat`

## Flujo de uso

1. Descargar la plantilla en blanco o cargar una plantilla mensual ya diligenciada.
2. Seleccionar SKU / Articulo.
3. Definir alcance con filtros dependientes del SKU.
4. Revisar informacion clave y visualizaciones.
5. Editar el forecast ajustado solicitado por mes.
6. Guardar ajustes y descargar soporte Excel.
