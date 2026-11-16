# Cómo usar el script de asignación de casos

Este script reasigna los casos de Pega según un archivo de Excel (la distribución del día). Lo hace solo, fila por fila. Vos tenés que iniciar sesión y estar atento a la terminal.

## Paso 0: Bajar el repositorio (solo la primera vez)

Necesitás acceso al repositorio `AlanGarber/AutoJP` en GitHub. Si no lo tenés, pedile acceso a Alan. Necesitás tener Git instalado.

1. Buscá `cmd` en el menú Inicio y abrilo.
2. Copiá y pegá estas dos líneas, una por una, apretando Enter después de cada una:

```
cd "C:\Users\SID\OneDrive - JPMorgan Chase\Desktop"
git clone https://github.com/AlanGarber/AutoJP.git
```

3. Si te pide iniciar sesión en GitHub, seguí los pasos en pantalla.

Se crea la carpeta `AutoJP` en tu Escritorio.

**Después, una sola vez:**

1. Dentro de `AutoJP\asignar-casos-empleados`, copiá el archivo `config_local.example.py` y renombrá la copia a `config_local.py`.
2. Abrí `config_local.py` con el Bloc de notas y reemplazá la URL por la que te pase Alan.
3. En la terminal, instalá las librerías con esta línea y apretá Enter:

```
pip install playwright openpyxl
```

Si alguna de estas líneas falla, pedí ayuda a Alan antes de seguir.

Si es la primera vez que lo usás en esta computadora, también podés seguir los pasos de instalación del `README.md`.

---

## Paso 1: Bajar el archivo de Excel

Bajá el archivo de distribución a tu computadora. Quedará en la carpeta de descargas.

Tiene que ser uno de estos archivos:

- `Distribution of Cases (Today).xlsx` (mañana)
- `Distribution.xlsx` (tarde)
- `Distribution Failed.xlsx` (failed)

## Paso 2: Arrastrarlo a la carpeta del script

Arrastrá el Excel a la carpeta `asignar-casos-empleados`, la misma donde está el archivo `ui_automation.py`.

## Paso 3: Abrir la terminal en la carpeta del script

Abrí la terminal (`cmd`) y pegá esta línea. Después apretá Enter:

```
cd "C:\Users\SID\OneDrive - JPMorgan Chase\Desktop\AutoJP\asignar-casos-empleados"
```

Con eso la terminal queda parada en la carpeta del script, donde arrastraste el Excel en el paso 2.

## Paso 4: Copiar y pegar el comando que corresponde

Elegí el comando según el nombre del archivo que bajaste:

| Archivo | Comando |
|---|---|
| `Distribution of Cases (Today).xlsx` (mañana) | `python ui_automation.py "Distribution of Cases (Today).xlsx"` |
| `Distribution.xlsx` (tarde) | `python ui_automation.py "Distribution.xlsx"` |
| `Distribution Failed.xlsx` (failed) | `python ui_automation.py "Distribution Failed.xlsx"` |

Copiá el comando completo, pegalo en la terminal y apretá Enter.

## Paso 5: Iniciar sesión

1. Se abre el navegador con la aplicación. El script hace solo los clicks de SSO.
2. Completá usuario, contraseña y MFA en el navegador.
3. Volvé a la terminal y apretá Enter para empezar.

## Paso 6: Esperar

- No toques el mouse ni el teclado mientras corre. El script trabaja solo.
- Procesa una fila por vez. Las filas de equipos que no sean AUTOCV ni FX ChaseTr se saltean solas.
- Al terminar, la terminal muestra el resumen.

## Paso 7: Revisar los resultados

En la carpeta `asignar-casos-empleados` quedan dos tipos de archivos:

- `resultados.csv`: una fila por cada caso, con el estado. Abrilo con Excel.
  - `OK`: el caso se reasignó.
  - `ERROR`: no se pudo reasignar. La columna de detalle dice en qué paso falló.
  - `OMITIDA`: el equipo no era AUTOCV ni FX ChaseTr.
  - `NO PROCESADA`: la corrida se detuvo antes de llegar a esa fila.
- `error_fila_N.png`: captura de pantalla de la fila que falló (N es el número de fila del Excel).

Antes de dar la tarea por terminada, verificá en Pega los casos que quedaron como `OK`.

---

## Si algo sale mal

- **Una fila dio ERROR:** el script sigue con la siguiente. Anotá los casos con error y revisalos a mano.
- **La terminal dice "no se pudo volver a la búsqueda":** la corrida se detuvo. Cerrá la terminal, revisá la pantalla de Pega y volvé a correr el script. Antes, borrá del Excel las filas que ya dieron `OK`, para no repetirlas.
- **Querés cortar la corrida:** apretá Ctrl+C en la terminal. El archivo `resultados.csv` se guarda igual con lo procesado hasta ese momento.

## Antes de empezar

- Probá primero con un Excel de una sola fila, con un caso que puedas revertir.
- No compartas el Excel ni el `resultados.csv` en chats o correos fuera del equipo, porque tienen números de caso.
