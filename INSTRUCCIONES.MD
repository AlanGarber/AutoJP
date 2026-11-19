# Instrucciones de las automatizaciones de AutoJP

Este archivo cubre **dos automatizaciones distintas**. Cada una tiene su propia carpeta, su propio Excel y su propio comando. Leé solo la sección de la que vas a usar.

| | Automatización A: Asignar casos de empleados | Automatización B: Reasignar casos por LOB |
|---|---|---|
| Qué hace | Reasigna los casos de la distribución del día a cada empleado | Reasigna cada caso a la work queue que corresponde a su LOB |
| Carpeta | `asignar-casos-empleados` | `reasignar-casos-pega-lob` |
| Excel de entrada | `Distribution of Cases (Today).xlsx` o `Distribution.xlsx` | El archivo de casos por LOB que te pasan (columnas A, B y C) |
| Comando | `python ui_automation.py "NOMBRE.xlsx"` | `python ui_automation.py "NOMBRE.xlsx"` |
| Estado | En uso | En prueba |

Las dos usan el mismo repositorio, así que el Paso 0 de la automatización A sirve también para la B.

---

# Automatización A: Asignar casos de empleados

Este script reasigna los casos de PRPC según un archivo de Excel (la distribución del día). Lo hace solo, fila por fila. Vos tenés que iniciar sesión y estar atento a la terminal.

## Uso diario (cuando ya instalaste todo)

1. Bajá el archivo de Excel de la distribución a tu computadora.
2. Arrastrá el Excel a la carpeta `AutoJP\asignar-casos-empleados`.
3. Abrí `devshell`, pegá la línea del paso 3 (con tu SID), apretá Enter.
4. Pegá el comando que corresponde al archivo (paso 4), apretá Enter, y seguí los pasos 5 y 6.

Si todavía no instalaste nada, empezá por el Paso 0.

### Paso 0: Bajar el repositorio y preparar la computadora (solo la primera vez)

Necesitás acceso al repositorio `AlanGarber/AutoJP` en GitHub. Si no lo tenés, pedile acceso a Alan. Necesitás tener Git instalado.

1. Buscá `devshell` en el menú Inicio y abrilo.
2. Antes de copiar, reemplazá `SID` por tu SID. Después copiá y pegá estas dos líneas, una por una, apretando Enter después de cada una:

```
cd "C:\Users\SID\OneDrive - JPMorgan Chase\Desktop"
git clone https://github.com/AlanGarber/AutoJP.git
```

3. Si te pide iniciar sesión en GitHub, seguí los pasos en pantalla.

Se crea la carpeta `AutoJP` en tu Escritorio.

**Después, una sola vez:**

1. En la devshell, copiá y pegá esta línea y apretá Enter. Esperá a que termine:

```
ds tool install python3.15
```

2. Dentro de `AutoJP\asignar-casos-empleados`, copiá el archivo `config_local.example.py` y renombrá la copia a `config_local.py`.
3. Abrí `config_local.py` con el Bloc de notas y reemplazá la URL por la de PRPC. **Guardá el archivo antes de cerrarlo** (menú Archivo > Guardar, o Ctrl+S).
4. En la devshell, instalá las librerías con esta línea y apretá Enter:

```
pip install playwright openpyxl
```

Si alguna de estas líneas falla, pedí ayuda a Alan antes de seguir.

Si es la primera vez que lo usás en esta computadora, también podés seguir los pasos de instalación del `README.md`.

---

### Paso 1: Bajar el archivo de Excel

Bajá el archivo de distribución a tu computadora. Quedará en la carpeta de descargas.

Tiene que ser uno de estos archivos:

- `Distribution of Cases (Today).xlsx` (mañana)
- `Distribution.xlsx` (tarde)
- `Distribution Failed.xlsx` (failed): **todavía no está implementada**. No la uses por ahora.

### Paso 2: Arrastrarlo a la carpeta del script

Arrastrá el Excel a la carpeta `asignar-casos-empleados`, la misma donde está el archivo `ui_automation.py`.

### Paso 3: Abrir la terminal en la carpeta del script

Abrí la terminal (`devshell`). Antes de copiar la línea, reemplazá `SID` por tu SID. Después pegala y apretá Enter:

```
cd "C:\Users\SID\OneDrive - JPMorgan Chase\Desktop\AutoJP\asignar-casos-empleados"
```

Con eso la terminal queda parada en la carpeta del script, donde arrastraste el Excel en el paso 2.

### Paso 4: Copiar y pegar el comando que corresponde

Elegí el comando según el nombre del archivo que bajaste:

| Archivo | Comando |
|---|---|
| `Distribution of Cases (Today).xlsx` (mañana) | `python ui_automation.py "Distribution of Cases (Today).xlsx"` |
| `Distribution.xlsx` (tarde) | `python ui_automation.py "Distribution.xlsx"` |
| `Distribution Failed.xlsx` (failed) | Todavía no está implementada |

Copiá el comando completo, pegalo en la terminal y apretá Enter.

### Paso 5: Iniciar sesión

1. Se abre el navegador con la aplicación. El script hace solo los clicks de SSO.
2. Completá usuario, contraseña y MFA en el navegador.
3. No hace falta apretar nada en la terminal: el script arranca solo cuando aparece la búsqueda de PRPC.

### Paso 6: Esperar

- No toques el mouse ni el teclado mientras corre. El script trabaja solo.
- Procesa una fila por vez. Las filas de equipos que no sean AUTOCV, FX ChaseTr ni iFX se saltean solas.
- Mirá la terminal mientras corre:
  - En verde: `OK - JPV ... asignado a ...`. Ese caso quedó reasignado.
  - En rojo: `ERROR - JPV ... no se pudo asignar a ...`, con el motivo. Avisale a Alan.
- Al terminar, la terminal muestra una tabla con los JPV que fallaron (y a quién estaban asignados) y el resumen de OK y errores. Si hubo errores, pasale esa tabla a Alan.
  - En amarillo: `JPV ... asignado al bot (B089707). Se saltea la fila.` No hace falta hacer nada: ese caso ya está con el bot y el script sigue con el siguiente.

---

### Si algo sale mal (Automatización A)

- **La terminal muestra un error en rojo:** avisale a Alan con el texto que dice la terminal. No vuelvas a correr el script hasta que te responda.
- **La terminal dice "no se pudo volver a la búsqueda":** la corrida se detuvo. Cerrá la terminal, avisale a Alan y no corras el script de nuevo hasta que te diga cómo seguir.
- **Querés cortar la corrida:** apretá Ctrl+C en la terminal y avisale a Alan.

### Antes de empezar (Automatización A)

- Probá primero con un Excel de una sola fila, con un caso que puedas revertir.
- No compartas el Excel en chats o correos fuera del equipo, porque tiene números de caso.

---

# Automatización B: Reasignar casos por LOB

Esta automatización reasigna cada caso de la basket de Pega a la work queue que corresponde a su LOB. Lee un Excel, fila por fila, y en cada una hace clic en el JPV, reasigna y guarda. Vos tenés que iniciar sesión y estar atento a la terminal.

> **Atención: esta automatización reasigna de verdad.** Cada fila que pasa las validaciones se guarda en Pega (Submit y Close). No hay vuelta atrás desde el script.

## Antes de empezar (Automatización B)

- Probá primero con un Excel de una sola fila, con un caso que puedas revertir a mano en Pega.
- El Excel tiene datos de casos y JPV. No lo compartas en chats o correos fuera del equipo.
- El Excel tiene que ser `.xlsx` y tener la LOB en la columna A (`Corp&FI`, `Retail` o `To check`), la columna B (solo para el caso especial, `FX Retail Inv`) y el JPV en la columna C (ID). Las demás columnas se ignoran.

## Primera vez (instalación)

Si ya hiciste el Paso 0 de la Automatización A, el repositorio ya está en tu computadora: no hace falta clonarlo de nuevo.

1. En la devshell, instalá las librerías con esta línea y apretá Enter:

```
pip install playwright openpyxl
```

2. Dentro de la carpeta `AutoJP\reasignar-casos-pega-lob`, copiá el archivo `config_local.example.py` y renombrá la copia a `config_local.py`.
3. Abrí `config_local.py` con el Bloc de notas. Reemplazá la URL por la de Pega y completá `XLSX_RUTA` con la ruta de la carpeta de red que te pasó Alan. **Guardá el archivo antes de cerrarlo** (menú Archivo > Guardar, o Ctrl+S).

Si alguna de estas líneas falla, pedí ayuda a Alan antes de seguir.

## Uso diario (Automatización B)

El Excel ya está en la carpeta de red del equipo: no hace falta descargarlo ni copiarlo. El script lo lee directo de ahí.

### Paso 1: Abrir la terminal en la carpeta del script

Abrí la terminal (`devshell`). Antes de copiar la línea, reemplazá `SID` por tu SID. Después pegala y apretá Enter:

```
cd "C:\Users\SID\OneDrive - JPMorgan Chase\Desktop\AutoJP\reasignar-casos-pega-lob"
```

### Paso 2: Correr el script

Copiá este comando, pegalo en la terminal y apretá Enter:

```
python ui_automation.py
```

El script toma el Excel de la carpeta de red, configurada en `config_local.py`. Si el Excel no está, lo avisa en la terminal y no abre el navegador.

### Paso 3: Iniciar sesión

1. Se abre el navegador con Pega. El script hace solo los clicks de SSO.
2. Completá usuario, contraseña y MFA en el navegador.
3. Cuando veas Pega, volvé a la terminal y apretá Enter.
4. El script abre la basket solo.

### Paso 4: Esperar

- No toques el mouse ni el teclado mientras corre. El script trabaja solo.
- Procesa una fila por vez. Para cada fila busca el JPV en la basket, lo abre, y reasigna según la LOB.
- Si el JPV no aparece en ninguna página de la basket, la fila se omite y sigue con la siguiente.
- Mirá la terminal mientras corre. Cada fila muestra un estado:

| Estado | Qué significa | Qué hacer |
|---|---|---|
| `OK` | El caso quedó reasignado a la work queue de su LOB | Nada |
| `OMITIDA` | No se tocó el caso: el case owner no es `FX1213`, el JPV no está en la basket, la LOB no tiene work queue o el JPV está vacío | Nada. El motivo queda en el detalle |
| `ERROR` | Algo falló al reasignar | Avisale a Alan con la línea de la terminal y el archivo `error_fila_N.png` |
| `NO PROCESADA` | La basket se cerró antes de llegar a la fila | Avisale a Alan |

### Paso 5: Revisar el resultado

Al terminar quedan dos cosas en la carpeta del script:

- `resultados.csv`: una fila por cada fila del Excel, con su estado y el detalle.
- `error_fila_N.png`: captura de la pantalla de Pega en la fila N, solo si esa fila dio `ERROR`.

Si hubo `ERROR` o `NO PROCESADA`, pasale `resultados.csv` y las capturas a Alan.

## Si algo sale mal (Automatización B)

- **"No encuentro el XLSX":** la carpeta de red no está conectada, o la ruta de `config_local.py` cambió. Revisá la conexión y avisale a Alan.
- **Hay una fila en `ERROR`:** avisale a Alan. No vuelvas a correr el Excel hasta que te responda, porque esa fila puede haber quedado a medio camino en Pega. Revisala a mano antes de volver a correr.
- **La basket se cerró:** las filas que faltaban quedan `NO PROCESADA`. Avisale a Alan antes de volver a correr.
- **Querés cortar la corrida:** apretá Ctrl+C en la terminal y avisale a Alan. Revisá en Pega la fila que estaba procesando.
