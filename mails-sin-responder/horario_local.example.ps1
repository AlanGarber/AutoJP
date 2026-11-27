# Copia este archivo como horario_local.ps1 (sin ".example") y completa las celdas.
# horario_local.ps1 NO se sube al repo: tiene nombres reales del equipo.
# Si existe, reemplaza la lectura del Excel de horarios (Back Ups.xlsx).
# Filas 4 a 7 y columnas D a H (lunes a viernes), tal como aparecen en el Excel.
# Actualizar a mano cuando cambie el horario.

$HorarioFijo = @{
    'D4' = 'NOMBRE'; 'E4' = 'NOMBRE'; 'F4' = 'NOMBRE'; 'G4' = 'NOMBRE'; 'H4' = 'NOMBRE'
    'D5' = 'NOMBRE'; 'E5' = 'NOMBRE'; 'F5' = 'NOMBRE'; 'G5' = 'NOMBRE'; 'H5' = 'NOMBRE'
    'D6' = 'NOMBRE'; 'E6' = 'NOMBRE'; 'F6' = 'NOMBRE'; 'G6' = 'NOMBRE'; 'H6' = 'NOMBRE'
    'D7' = 'NOMBRE'; 'E7' = 'NOMBRE'; 'F7' = 'NOMBRE'; 'G7' = 'NOMBRE'; 'H7' = 'NOMBRE'
}
