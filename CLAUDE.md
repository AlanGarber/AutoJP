# AutoJP: instrucciones para Claude

Repo de automatizaciones de UI para el flujo diario del equipo. Todo el contenido, comentarios y documentación están en español (rioplatense, con voseo). Los nombres de variables y funciones de Python van en inglés, como en el script actual.

Antes de tocar cualquier automatización, leer `CONOCIMIENTO.md`: contiene las decisiones tomadas, el estado de cada automatización y los pendientes.

## Estructura

- `README.md` (raíz): índice de automatizaciones. Hay que sumar cada nueva a su tabla.
- `<nombre-en-minusculas-con-guiones>/`: una carpeta por automatización, con:
  - `README.md`: instalación, configuración, archivo de entrada y ejecución, en pasos simples.
  - `ui_automation.py` (u otro script): un solo archivo, con una sección `CONFIGURACIÓN` arriba de todo.
  - `config_local.example.py`: plantilla para la URL real (la copia local `config_local.py` no se sube).

## Stack

- Python 3.9+ con Playwright (API sincrónica), `openpyxl` para leer los XLSX y el módulo `csv` para el resultado.
- Los archivos de entrada son siempre `.xlsx`.
- No agregues dependencias nuevas sin preguntar.
- Los selectores del script son placeholders hasta que el usuario los complete con codegen o con las DevTools.

## Reglas de trabajo

- **Commits y push:** solo cuando el usuario lo pida explícitamente. Mensajes de commit en español, cortos y descriptivos.
- **No correr contra la app real** sin confirmación. El script pide login SSO y un Enter manual, así que el usuario lo ejecuta en su propia terminal.
- **Verificación mínima:** antes de dar un cambio por listo, correr `python -m py_compile <script>`. Si hay Playwright instalado, probar con un XLSX de prueba local.
- **Datos sensibles:** los XLSX tienen assignees y JPV. No pegarlos en commits, issues ni respuestas a servicios externos. Los XLSX (`*.xlsx`), los archivos generados (`resultados.csv`, `error_fila_*.png`) y `config_local.py` están en `.gitignore` y no se suben.
- **Credenciales:** nunca guardar usuarios, contraseñas, tokens ni cookies en el repo.

## Cómo agregar una automatización nueva

1. Crear la carpeta con nombre descriptivo en minúsculas y guiones.
2. Escribir el script con la sección `CONFIGURACIÓN` arriba, y los selectores como placeholders.
3. Escribir su `README.md` siguiendo el formato de `asignar-casos-empleados/README.md`.
4. Documentar en el README el formato del XLSX con una tabla, sin subir archivos con datos.
5. Sumarla a la tabla del `README.md` raíz.

## Cómo pedir cosas al usuario

- Si falta un dato (URL, selector, nombre de columna), preguntarlo en una sola pregunta concreta.
- Si una decisión es difícil de revertir (push, borrar archivos, cambiar el formato de entrada), confirmarla antes.
