# Cómo usar el script de asignación de casos

Este script reasigna los casos de PRPC según un archivo de Excel (la distribución del día). Lo hace solo, fila por fila. Vos tenés que iniciar sesión y estar atento a la terminal.

## Uso diario (cuando ya instalaste todo)

1. Bajá el archivo de Excel de la distribución a tu computadora.
2. Arrastrá el Excel a la carpeta `AutoJP\asignar-casos-empleados`.
3. Abrí `devshell`, pegá la línea del paso 3 (con tu SID), apretá Enter.
4. Pegá el comando que corresponde al archivo (paso 4), apretá Enter, y seguí los pasos 5 y 6.

Si todavía no instalaste nada, empezá por el Paso 0.

## Paso 0: Bajar el repositorio y preparar la computadora (solo la primera vez)

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

## Paso 1: Bajar el archivo de Excel

Bajá el archivo de distribución a tu computadora. Quedará en la carpeta de descargas.

Tiene que ser uno de estos archivos:

- `Distribution of Cases (Today).xlsx` (mañana)
- `Distribution.xlsx` (tarde)
- `Distribution Failed.xlsx` (failed): **todavía no está implementada**. No la uses por ahora.

## Paso 2: Arrastrarlo a la carpeta del script

Arrastrá el Excel a la carpeta `asignar-casos-empleados`, la misma donde está el archivo `ui_automation.py`.

## Paso 3: Abrir la terminal en la carpeta del script

Abrí la terminal (`devshell`). Antes de copiar la línea, reemplazá `SID` por tu SID. Después pegala y apretá Enter:

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
| `Distribution Failed.xlsx` (failed) | Todavía no está implementada |

Copiá el comando completo, pegalo en la terminal y apretá Enter.

## Paso 5: Iniciar sesión

1. Se abre el navegador con la aplicación. El script hace solo los clicks de SSO.
2. Completá usuario, contraseña y MFA en el navegador.
3. Volvé a la terminal y apretá Enter para empezar.

## Paso 6: Esperar

- No toques el mouse ni el teclado mientras corre. El script trabaja solo.
- Procesa una fila por vez. Las filas de equipos que no sean AUTOCV ni FX ChaseTr se saltean solas.
- Mirá la terminal mientras corre:
  - En verde: `OK - JPV ... asignado a ...`. Ese caso quedó reasignado.
  - En rojo: `ERROR - JPV ... no se pudo asignar a ...`, con el motivo. Avisale a Alan.
  - En amarillo: `JPV ... asignado al bot (B089707). Se saltea la fila.` No hace falta hacer nada: ese caso ya está con el bot y el script sigue con el siguiente.

---

## Si algo sale mal

- **La terminal muestra un error en rojo:** avisale a Alan con el texto que dice la terminal. No vuelvas a correr el script hasta que te responda.
- **La terminal dice "no se pudo volver a la búsqueda":** la corrida se detuvo. Cerrá la terminal, avisale a Alan y no corras el script de nuevo hasta que te diga cómo seguir.
- **Querés cortar la corrida:** apretá Ctrl+C en la terminal y avisale a Alan.

## Antes de empezar

- Probá primero con un Excel de una sola fila, con un caso que puedas revertir.
- No compartas el Excel en chats o correos fuera del equipo, porque tiene números de caso.
