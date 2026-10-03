# Automatización BN

Aplicación de escritorio para cargar marcaciones desde Excel/TXT y registrarlas en **Maestría de Asistencia**.

## Funciones
- Abrir Excel (`.xlsx`, `.xlsm`) o TXT/CSV.
- Previsualizar y normalizar la data.
- Agregar, modificar y eliminar filas.
- Validar código, fecha y hora antes de automatizar.
- Ejecutar la automatización en segundo plano.
- Ver el estado de cada registro en la interfaz.
- Cancelar el proceso.
- Exportar el resultado a Excel.

## Columnas esperadas
El programa reconoce nombres parecidos a:
- `Cod. Emplea` / `Código` / `Codigo`
- `Incidente`
- `Fecha inicio`
- `Hora inicio`
- `Hora fin`

Regla:
- Si `Incidente` contiene `INGRESO`, usa `Hora inicio`.
- Si `Incidente` contiene `SALIDA`, usa `Hora fin`.

## Importante sobre GitHub Codespaces
Codespaces normalmente corre Linux. **PyInstaller no genera un .exe de Windows desde Linux**.
Por eso este proyecto incluye un workflow de GitHub Actions que compila el `.exe` usando un runner Windows.

### En Codespaces
```bash
pip install -r requirements.txt
python app.py
```
La interfaz gráfica no es útil dentro de Codespaces porque la automatización debe ejecutarse en el mismo Windows donde está abierto Maestría de Asistencia. Usa Codespaces para editar/probar lógica y GitHub Actions para generar el EXE.

### Generar el EXE con GitHub Actions
1. Sube estos archivos a tu repositorio.
2. Ve a `Actions`.
3. Ejecuta `Build Windows EXE`.
4. Al finalizar, descarga el artifact `AUTOMATIZACION_BN-Windows`.

## Primer uso en Windows
1. Abre **Maestría de Asistencia** e inicia sesión normalmente.
2. Ejecuta `AUTOMATIZACION_BN.exe`.
3. Carga el Excel/TXT.
4. Revisa/modifica la tabla.
5. Pulsa `Iniciar automatización`.

> La automatización usa controles de Windows cuando están disponibles y coordenadas relativas de respaldo para aplicaciones antiguas. Si la interfaz de Maestría de Asistencia cambia, ajusta `automation_profile.json`.
