<#
Reporte de mails recibidos en el día que no fueron respondidos, en una o
varias casillas compartidas (y opcionalmente en tu casilla personal).

Por cada casilla lee la Bandeja de entrada (y sus subcarpetas) y Elementos
enviados del Outlook de escritorio, y marca como "sin responder" cada mail
recibido en el día que nadie del equipo respondió.

Usa solo PowerShell (viene con Windows) y el Outlook instalado:
no hace falta instalar nada.

Uso:
    powershell -NoProfile -ExecutionPolicy Bypass -File .\reporte_no_respondidos.ps1
    powershell -NoProfile -ExecutionPolicy Bypass -File .\reporte_no_respondidos.ps1 -Fecha 2026-11-12

Genera, en esta carpeta:
    reporte_no_respondidos_AAAA-MM-DD.csv
    reporte_no_respondidos_AAAA-MM-DD.html
#>

param(
    # Día a analizar. Por defecto, hoy.
    [datetime]$Fecha = (Get-Date).Date
)

# =====================================================================
# CONFIGURACIÓN
# =====================================================================

# Casillas a revisar: el email o el nombre tal como aparece en Outlook.
# 'PERSONAL' agrega tu propia casilla.
$Casillas = @(
    'casilla.compartida1@empresa.com'
    'casilla.compartida2@empresa.com'
)

# Emails de las personas que responden desde las casillas. Si alguien del equipo
# contesta con la casilla en copia, esa respuesta cuenta. Además, los mails que
# se mandan entre ustedes no aparecen como pendientes. Podés dejarla vacía: @().
$Equipo = @(
    # 'persona1@empresa.com'
    # 'persona2@empresa.com'
)

# Revisar también las subcarpetas de la Bandeja de entrada (por si hay reglas que mueven mails).
$IncluirSubcarpetas = $true

# $true: ignora los mails donde la casilla está solo en CC/CCO o llegó por una lista de distribución.
$SoloSiEstaEnPara = $false

# $true: si alguien reenvió el mail (FW/RV), cuenta como respondido.
$ContarReenvioComoRespuesta = $false

# $true: muestra una sola fila por conversación (el último mail sin responder del hilo).
$AgruparPorConversacion = $true

# Remitentes que nunca esperan respuesta. Se busca el texto dentro del email del remitente.
$RemitentesIgnorados = @(
    'noreply', 'no-reply', 'donotreply', 'do-not-reply', 'mailer-daemon', 'postmaster'
)

# Asuntos que se ignoran (expresiones regulares, sin distinguir mayúsculas).
$AsuntosIgnorados = @(
    '^Automatic reply', '^Respuesta automática', '^Out of Office', '^Fuera de la oficina',
    '^Undeliverable', '^No se puede entregar', '^Accepted:', '^Aceptado:', '^Declined:', '^Rechazado:'
)

# Dónde se guardan los reportes. Por defecto, la carpeta del script.
$CarpetaSalida = $PSScriptRoot

# Abrir el reporte HTML en el navegador al terminar.
$AbrirReporteAlTerminar = $true

# =====================================================================

# Lista local del equipo (equipo_local.ps1, fuera del repo). Si existe, reemplaza $Equipo.
$RutaEquipoLocal = if ($PSScriptRoot) { Join-Path $PSScriptRoot 'equipo_local.ps1' } else { Join-Path (Get-Location) 'equipo_local.ps1' }
if (Test-Path $RutaEquipoLocal) { . $RutaEquipoLocal }
# Lista de casillas (casillas_local.ps1, fuera del repo). Si existe, reemplaza $Casillas.
$RutaCasillasLocal = if ($PSScriptRoot) { Join-Path $PSScriptRoot 'casillas_local.ps1' } else { Join-Path (Get-Location) 'casillas_local.ps1' }
if (Test-Path $RutaCasillasLocal) { . $RutaCasillasLocal }

$ErrorActionPreference = 'Stop'

# Constantes de Outlook
$olFolderSentMail = 5
$olFolderInbox = 6
$olMail = 43
$PropUltimaAccion = 'http://schemas.microsoft.com/mapi/proptag/0x10810003'      # PR_LAST_VERB_EXECUTED
$PropUltimaAccionHora = 'http://schemas.microsoft.com/mapi/proptag/0x10820040'  # PR_LAST_VERB_EXECUTION_TIME
$PropSmtp = 'http://schemas.microsoft.com/mapi/proptag/0x39FE001E'              # PR_SMTP_ADDRESS
$AccionResponder = 102
$AccionResponderATodos = 103
$AccionReenviar = 104
$PatronReenvio = '^\s*(FW|FWD|RV|TR|WG)\s*:'


function New-AddressSet {
    return , (New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase))
}


function Get-MyAddresses($namespace) {
    $set = New-AddressSet
    foreach ($account in $namespace.Accounts) {
        if ($account.SmtpAddress) { [void]$set.Add($account.SmtpAddress) }
    }
    try {
        $exUser = $namespace.CurrentUser.AddressEntry.GetExchangeUser()
        if ($exUser) { [void]$set.Add($exUser.PrimarySmtpAddress) }
    } catch {}
    if ($namespace.CurrentUser.Name) { [void]$set.Add($namespace.CurrentUser.Name) }
    return , $set
}


function Open-Mailbox($namespace, $name, $me) {
    # Devuelve la Bandeja de entrada, los Enviados y las direcciones de la casilla.
    if ($name -eq 'PERSONAL') {
        return [pscustomobject]@{
            Nombre      = 'Personal'
            Direcciones = $me
            Inbox       = $namespace.GetDefaultFolder($olFolderInbox)
            Enviados    = $namespace.GetDefaultFolder($olFolderSentMail)
        }
    }

    $recipient = $namespace.CreateRecipient($name)
    if (-not $recipient.Resolve()) {
        throw "No encontré '$name' en la libreta de direcciones. Revisá el nombre o el email."
    }
    $addresses = New-AddressSet
    [void]$addresses.Add($name)
    [void]$addresses.Add($recipient.Name)
    try {
        $exUser = $recipient.AddressEntry.GetExchangeUser()
        if ($exUser) { [void]$addresses.Add($exUser.PrimarySmtpAddress); [void]$addresses.Add($exUser.Name) }
    } catch {}

    # 1) Si la casilla ya aparece en el panel de carpetas de Outlook, se usa esa:
    #    así se ven también las subcarpetas.
    $inbox = $null
    $sent = $null
    foreach ($store in $namespace.Stores) {
        $storeName = $null
        try { $storeName = [string]$store.DisplayName } catch { continue }
        if (-not $addresses.Contains($storeName)) { continue }
        try { $inbox = $store.GetDefaultFolder($olFolderInbox) } catch {}
        try { $sent = $store.GetDefaultFolder($olFolderSentMail) } catch {}
        break
    }

    # 2) Si no, se abre como carpeta compartida.
    if (-not $inbox) { $inbox = $namespace.GetSharedDefaultFolder($recipient, $olFolderInbox) }
    if (-not $sent) {
        try { $sent = $namespace.GetSharedDefaultFolder($recipient, $olFolderSentMail) }
        catch { Write-Warning "No pude abrir los Enviados de '$name'. Se usa solo la marca de 'Respondido' de cada mail." }
    }

    return [pscustomobject]@{
        Nombre      = [string]$recipient.Name
        Direcciones = $addresses
        Inbox       = $inbox
        Enviados    = $sent
    }
}


function Get-SenderSmtp($mail) {
    try {
        if ($mail.SenderEmailType -eq 'EX') {
            $exUser = $mail.Sender.GetExchangeUser()
            if ($exUser) { return $exUser.PrimarySmtpAddress }
        }
    } catch {}
    return [string]$mail.SenderEmailAddress
}


function Get-RecipientSmtp($recipient) {
    try { return [string]$recipient.PropertyAccessor.GetProperty($PropSmtp) } catch {}
    return [string]$recipient.Address
}


function Get-OwnerRole($mail, $owner) {
    # En qué campo está la casilla: Para, CC, CCO o Lista/otro (p. ej. lista de distribución).
    $role = $null
    foreach ($recipient in $mail.Recipients) {
        $isOwner = $owner.Contains((Get-RecipientSmtp $recipient)) -or $owner.Contains([string]$recipient.Name)
        if (-not $isOwner) { continue }
        switch ($recipient.Type) {
            1 { return 'Para' }
            2 { $role = 'CC' }
            3 { if (-not $role) { $role = 'CCO' } }
        }
    }
    if ($role) { return $role }
    return 'Lista/otro'
}


function Get-Prop($mail, $prop) {
    try { return $mail.PropertyAccessor.GetProperty($prop) } catch { return $null }
}


function Get-ConversationKey($mail) {
    try { if ($mail.ConversationID) { return [string]$mail.ConversationID } } catch {}
    try { return 'tema:' + [string]$mail.ConversationTopic } catch { return 'tema:' + [string]$mail.Subject }
}


function Convert-Received($mail, $folderName, $owner) {
    $lastVerbTime = $null
    $raw = Get-Prop $mail $PropUltimaAccionHora
    if ($raw) {
        try { $lastVerbTime = $mail.PropertyAccessor.UTCToLocalTime($raw) } catch {}
    }
    [pscustomobject]@{
        Recibido         = [datetime]$mail.ReceivedTime
        De               = [string]$mail.SenderName
        Email            = Get-SenderSmtp $mail
        Asunto           = [string]$mail.Subject
        Carpeta          = $folderName
        EnCampo          = Get-OwnerRole $mail $owner
        ConversationId   = Get-ConversationKey $mail
        MessageClass     = [string]$mail.MessageClass
        UltimaAccion     = Get-Prop $mail $PropUltimaAccion
        UltimaAccionHora = $lastVerbTime
    }
}


function Convert-Sent($mail, $folderName) {
    [pscustomobject]@{
        Enviado        = [datetime]$mail.SentOn
        Asunto         = [string]$mail.Subject
        ConversationId = Get-ConversationKey $mail
    }
}


function Read-Folder($folder, $timeField, $from, $to, $recursive, $converter) {
    # Recorre la carpeta de más nuevo a más viejo y corta al pasar $from,
    # así no lee buzones enteros.
    $out = New-Object System.Collections.Generic.List[object]
    $items = $folder.Items
    $items.Sort("[$timeField]", $true)
    $item = $items.GetFirst()
    while ($null -ne $item) {
        $time = $null
        try { $time = $item.$timeField } catch {}
        if ($null -ne $time) {
            if ($time -lt $from) { break }
            if ($time -lt $to -and $item.Class -eq $olMail) {
                try { $out.Add((& $converter $item $folder.Name)) }
                catch { Write-Warning "No pude leer un mail de '$($folder.Name)': $($_.Exception.Message)" }
            }
        }
        $item = $items.GetNext()
    }
    if ($recursive) {
        foreach ($sub in $folder.Folders) {
            $subItems = Read-Folder $sub $timeField $from $to $true $converter
            foreach ($x in $subItems) { $out.Add($x) }
        }
    }
    return , $out
}


function Find-Unanswered($received, $sent, $team, $dayEnd, $mailboxName) {
    # $team: direcciones de la casilla, tuyas y del equipo. Sus mails no se
    # reportan y, si están en la misma conversación, cuentan como respuesta.
    $replies = New-Object System.Collections.Generic.List[object]
    foreach ($s in $sent) { $replies.Add($s) }
    foreach ($r in $received) {
        if ($team.Contains($r.Email)) {
            $replies.Add([pscustomobject]@{ Enviado = $r.Recibido; Asunto = $r.Asunto; ConversationId = $r.ConversationId })
        }
    }

    $repliesByConversation = @{}
    foreach ($s in $replies) {
        if (-not $ContarReenvioComoRespuesta -and $s.Asunto -match $PatronReenvio) { continue }
        if (-not $repliesByConversation.ContainsKey($s.ConversationId)) {
            $repliesByConversation[$s.ConversationId] = New-Object System.Collections.Generic.List[object]
        }
        $repliesByConversation[$s.ConversationId].Add($s)
    }

    $replyVerbs = @($AccionResponder, $AccionResponderATodos)
    if ($ContarReenvioComoRespuesta) { $replyVerbs += $AccionReenviar }

    $stats = [pscustomobject]@{ Casilla = $mailboxName; Recibidos = 0; Ignorados = 0; Respondidos = 0; SinResponder = 0 }
    $pending = New-Object System.Collections.Generic.List[object]

    foreach ($r in $received) {
        $stats.Recibidos++

        $ignore = $team.Contains($r.Email) -or
            ($r.MessageClass -notlike 'IPM.Note*') -or
            ($SoloSiEstaEnPara -and $r.EnCampo -ne 'Para')
        foreach ($p in $RemitentesIgnorados) { if ($r.Email -like "*$p*") { $ignore = $true } }
        foreach ($p in $AsuntosIgnorados) { if ($r.Asunto -match $p) { $ignore = $true } }
        if ($ignore) { $stats.Ignorados++; continue }

        # 1) Outlook marca el mail como respondido, sin importar quién del equipo respondió.
        $answered = ($replyVerbs -contains $r.UltimaAccion) -and
            ($null -eq $r.UltimaAccionHora -or ($r.UltimaAccionHora -ge $r.Recibido -and $r.UltimaAccionHora -lt $dayEnd))

        # 2) O hay un mail enviado (o del equipo) en la misma conversación, posterior.
        if (-not $answered -and $repliesByConversation.ContainsKey($r.ConversationId)) {
            foreach ($s in $repliesByConversation[$r.ConversationId]) {
                if ($s.Enviado -gt $r.Recibido -and $s.Enviado -lt $dayEnd) { $answered = $true; break }
            }
        }

        if ($answered) { $stats.Respondidos++ } else { $pending.Add($r) }
    }

    $now = Get-Date
    $reference = if ($now -lt $dayEnd) { $now } else { $dayEnd }

    $rows = @()
    if ($AgruparPorConversacion) {
        foreach ($group in ($pending | Group-Object ConversationId)) {
            $latest = $group.Group | Sort-Object Recibido -Descending | Select-Object -First 1
            $rows += New-ReportRow $latest $group.Count $reference $mailboxName
        }
    } else {
        foreach ($r in $pending) { $rows += New-ReportRow $r 1 $reference $mailboxName }
    }
    $stats.SinResponder = $pending.Count

    return [pscustomobject]@{
        Filas        = @($rows | Sort-Object Recibido)
        Estadisticas = $stats
    }
}


function New-ReportRow($r, $countInThread, $reference, $mailboxName) {
    [pscustomobject]@{
        Casilla           = $mailboxName
        Recibido          = $r.Recibido.ToString('yyyy-MM-dd HH:mm')
        De                = $r.De
        Email             = $r.Email
        Asunto            = $r.Asunto
        Carpeta           = $r.Carpeta
        EnCampo           = $r.EnCampo
        HorasSinRespuesta = [math]::Round(($reference - $r.Recibido).TotalHours, 1)
        MailsSinResponder = $countInThread
    }
}


function HtmlText($texto) {
    return [System.Net.WebUtility]::HtmlEncode([string]$texto)
}


function Get-CampoClave($enCampo) {
    # Agrupa el campo de Outlook en tres filtros: para nosotros, en copia y lista/otro.
    switch ($enCampo) {
        'Para'  { return 'para' }
        'CC'    { return 'copia' }
        'CCO'   { return 'copia' }
        default { return 'otro' }
    }
}


function Get-ClaseHoras($horas) {
    # Color según cuánto tiempo lleva sin respuesta: verde < 1 h, amarillo 1-4 h, rojo >= 4 h.
    if ($horas -ge 4) { return 'horas-alta' }
    if ($horas -ge 1) { return 'horas-media' }
    return 'horas-baja'
}


function New-ResumenHtml($stats) {
    $suma = { param($propiedad) [int](@($stats | ForEach-Object { [int]$_.$propiedad }) | Measure-Object -Sum).Sum }
    $totales = [pscustomobject]@{
        Casilla     = 'Total'
        Recibidos   = & $suma 'Recibidos'
        Ignorados   = & $suma 'Ignorados'
        Respondidos = & $suma 'Respondidos'
        SinResponder = & $suma 'SinResponder'
    }

    $tarjetas = New-Object System.Text.StringBuilder
    foreach ($s in (@($stats) + @($totales))) {
        $clase = if ($s.Casilla -eq 'Total') { 'tarjeta total' } else { 'tarjeta' }
        [void]$tarjetas.Append(
            "<div class=`"$clase`">" +
            "<div class=`"tarjeta-titulo`">$(HtmlText $s.Casilla)</div>" +
            "<div class=`"metricas`">" +
            "<div><span class=`"num`">$($s.Recibidos)</span><span class=`"lbl`">recibidos</span></div>" +
            "<div><span class=`"num`">$($s.Respondidos)</span><span class=`"lbl`">respondidos</span></div>" +
            "<div class=`"destacado`"><span class=`"num`">$($s.SinResponder)</span><span class=`"lbl`">sin responder</span></div>" +
            "</div>" +
            "<div class=`"nota`">$($s.Ignorados) ignorados</div>" +
            "</div>")
    }
    return '<section class="tarjetas">' + $tarjetas.ToString() + '</section>'
}


function New-DetalleHtml($rows, $stats) {
    $casillas = @($stats | ForEach-Object { [string]$_.Casilla })
    $carpetas = @($rows | ForEach-Object { [string]$_.Carpeta } | Sort-Object -Unique)

    $botonesCampo = '<button data-campo-btn="todos" class="activo">Todos</button>' +
        '<button data-campo-btn="para">Para nosotros</button>' +
        '<button data-campo-btn="copia">En copia</button>' +
        '<button data-campo-btn="otro">Lista/otro</button>'

    $botonesCasilla = '<button data-casilla-btn="todas" class="activo">Todas</button>'
    for ($i = 0; $i -lt $casillas.Count; $i++) {
        $botonesCasilla += "<button data-casilla-btn=`"$i`">$(HtmlText $casillas[$i])</button>"
    }

    $botonesCarpeta = '<button data-carpeta-btn="todas" class="activo">Todas</button>'
    for ($i = 0; $i -lt $carpetas.Count; $i++) {
        $botonesCarpeta += "<button data-carpeta-btn=`"$i`">$(HtmlText $carpetas[$i])</button>"
    }

    $filtros = '<div class="filtros">' +
        '<div class="grupo"><span class="etiqueta">Campo</span>' + $botonesCampo + '</div>' +
        '<div class="grupo"><span class="etiqueta">Casilla</span>' + $botonesCasilla + '</div>' +
        '<div class="grupo"><span class="etiqueta">Carpeta</span>' + $botonesCarpeta + '</div>' +
        '</div>'

    $filas = New-Object System.Text.StringBuilder
    foreach ($r in $rows) {
        $idx = [array]::IndexOf($casillas, [string]$r.Casilla)
        $idxCarpeta = [array]::IndexOf($carpetas, [string]$r.Carpeta)
        $campo = Get-CampoClave $r.EnCampo
        $celdas = @(
            "<td>$(HtmlText $r.Casilla)</td>",
            "<td class=`"nowrap`">$(HtmlText $r.Recibido)</td>",
            "<td>$(HtmlText $r.De)</td>",
            "<td class=`"email`">$(HtmlText $r.Email)</td>",
            "<td class=`"asunto`">$(HtmlText $r.Asunto)</td>",
            "<td>$(HtmlText $r.Carpeta)</td>",
            "<td><span class=`"badge badge-$campo`">$(HtmlText $r.EnCampo)</span></td>",
            "<td class=`"num-col`"><span class=`"horas $(Get-ClaseHoras $r.HorasSinRespuesta)`">$(HtmlText $r.HorasSinRespuesta) h</span></td>",
            "<td class=`"num-col`">$(HtmlText $r.MailsSinResponder)</td>"
        )
        [void]$filas.Append(
            "<tr data-campo=`"$campo`" data-casilla=`"$idx`" data-carpeta=`"$idxCarpeta`">" + ($celdas -join '') + "</tr>")
    }

    $thead = '<tr><th>Casilla</th><th>Recibido</th><th>De</th><th>Email</th><th>Asunto</th><th>Carpeta</th><th>En campo</th><th>Horas sin respuesta</th><th>Mails del hilo</th></tr>'

    $leyenda = '<p class="leyenda">Horas sin respuesta: <span class="horas horas-baja">menos de 1 h</span> ' +
        '<span class="horas horas-media">1 a 4 h</span> <span class="horas horas-alta">4 h o más</span></p>'

    $jsFiltros = @'
<script>
  var estado = { campo: 'todos', casilla: 'todas', carpeta: 'todas' };
  function aplicarFiltros() {
    var visibles = 0;
    document.querySelectorAll('#detalle tbody tr').forEach(function (tr) {
      var ok = (estado.campo === 'todos' || tr.dataset.campo === estado.campo) &&
               (estado.casilla === 'todas' || tr.dataset.casilla === estado.casilla) &&
               (estado.carpeta === 'todas' || tr.dataset.carpeta === estado.carpeta);
      tr.style.display = ok ? '' : 'none';
      if (ok) {
        visibles++;
        tr.classList.toggle('par', visibles % 2 === 0);
      }
    });
    document.querySelectorAll('[data-campo-btn]').forEach(function (b) {
      b.classList.toggle('activo', b.dataset.campoBtn === estado.campo);
    });
    document.querySelectorAll('[data-casilla-btn]').forEach(function (b) {
      b.classList.toggle('activo', b.dataset.casillaBtn === estado.casilla);
    });
    document.querySelectorAll('[data-carpeta-btn]').forEach(function (b) {
      b.classList.toggle('activo', b.dataset.carpetaBtn === estado.carpeta);
    });
    document.getElementById('contador').textContent = visibles + (visibles === 1 ? ' mail visible' : ' mails visibles');
  }
  document.querySelectorAll('[data-campo-btn]').forEach(function (b) {
    b.addEventListener('click', function () { estado.campo = b.dataset.campoBtn; aplicarFiltros(); });
  });
  document.querySelectorAll('[data-casilla-btn]').forEach(function (b) {
    b.addEventListener('click', function () { estado.casilla = b.dataset.casillaBtn; aplicarFiltros(); });
  });
  document.querySelectorAll('[data-carpeta-btn]').forEach(function (b) {
    b.addEventListener('click', function () { estado.carpeta = b.dataset.carpetaBtn; aplicarFiltros(); });
  });
  aplicarFiltros();
</script>
'@

    return '<h2>Detalle</h2>' + $filtros +
        '<p id="contador" class="contador"></p>' +
        $leyenda +
        '<div class="tabla-wrap"><table id="detalle"><thead>' + $thead + '</thead><tbody>' +
        $filas.ToString() + '</tbody></table></div>' +
        $jsFiltros
}


function Write-Report($rows, $stats, $day, $folder) {
    $stamp = $day.ToString('yyyy-MM-dd')
    $csvPath = Join-Path $folder "reporte_no_respondidos_$stamp.csv"
    $htmlPath = Join-Path $folder "reporte_no_respondidos_$stamp.html"
    $rows = @($rows)

    # -UseCulture usa el separador de listas de Windows, así Excel lo abre en columnas.
    $rows | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -UseCulture

    $head = @'
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
  /* Colores de marca: cambiarlos acá. Acento #117ACA y navy #0C2A3A, según fuentes públicas. */
  :root { --navy: #0C2A3A; --azul: #117ACA; --texto: #1d2733; --gris: #5b6776; --borde: #e3e8ef; --fondo: #f4f6f9; }
  * { box-sizing: border-box; }
  body { margin: 0; font-family: "Segoe UI", Roboto, Arial, sans-serif; font-size: 15px; line-height: 1.45; color: var(--texto); background: var(--fondo); }
  header.encabezado { background: var(--navy); color: #fff; padding: 22px 32px; }
  header.encabezado h1 { margin: 0; font-size: 24px; font-weight: 600; }
  header.encabezado .fecha { margin: 4px 0 0; color: #cfe0f1; font-size: 14px; }
  main { padding: 24px 32px 40px; }
  h2 { font-size: 17px; font-weight: 600; margin: 28px 0 12px; }
  .tarjetas { display: grid; grid-template-columns: repeat(auto-fit, minmax(230px, 1fr)); gap: 16px; }
  .tarjeta { background: #fff; border: 1px solid var(--borde); border-radius: 10px; padding: 16px 18px; box-shadow: 0 1px 2px rgba(0,0,0,.04); }
  .tarjeta.total { border-color: var(--azul); }
  .tarjeta-titulo { font-weight: 600; margin-bottom: 10px; word-break: break-word; }
  .metricas { display: flex; gap: 20px; flex-wrap: wrap; }
  .metricas > div { display: flex; flex-direction: column; }
  .num { font-size: 24px; font-weight: 700; line-height: 1.1; }
  .lbl { font-size: 12px; color: var(--gris); text-transform: uppercase; letter-spacing: .03em; }
  .destacado .num { color: #b42318; }
  .nota { margin-top: 10px; font-size: 12px; color: var(--gris); }
  .filtros { background: #fff; border: 1px solid var(--borde); border-radius: 10px; padding: 10px 16px; }
  .grupo { display: flex; flex-wrap: wrap; align-items: center; gap: 8px; margin: 8px 0; }
  .etiqueta { width: 72px; font-size: 12px; font-weight: 600; color: var(--gris); text-transform: uppercase; letter-spacing: .03em; }
  .filtros button { padding: 6px 13px; border: 1px solid #c9d3df; background: #fff; color: var(--texto); border-radius: 999px; cursor: pointer; font: inherit; font-size: 13px; }
  .filtros button:hover { border-color: var(--azul); }
  .filtros button.activo { background: var(--azul); border-color: var(--azul); color: #fff; }
  .contador { margin: 12px 2px 4px; color: var(--gris); font-size: 13px; }
  .leyenda { margin: 4px 2px 10px; color: var(--gris); font-size: 13px; }
  .leyenda .horas { margin-right: 4px; }
  .tabla-wrap { overflow: auto; max-height: 68vh; border: 1px solid var(--borde); border-radius: 10px; background: #fff; }
  table { border-collapse: collapse; width: 100%; font-size: 14px; }
  thead th { position: sticky; top: 0; background: #eaf4fb; text-align: left; font-weight: 600; padding: 11px 12px; border-bottom: 1px solid var(--borde); white-space: nowrap; z-index: 1; }
  td { padding: 10px 12px; border-bottom: 1px solid var(--borde); vertical-align: top; }
  tr.par td { background: #fafbfd; }
  tbody tr:hover td { background: #eaf2fb; }
  .nowrap { white-space: nowrap; }
  .email { color: var(--gris); word-break: break-all; }
  .asunto { min-width: 260px; }
  .num-col { text-align: right; white-space: nowrap; }
  .badge { display: inline-block; padding: 2px 10px; border-radius: 999px; font-size: 12px; font-weight: 600; white-space: nowrap; }
  .badge-para { background: #dcfce7; color: #166534; }
  .badge-copia { background: #e0ecff; color: #1e40af; }
  .badge-otro { background: #f1f1f4; color: #4b5563; }
  .horas { display: inline-block; padding: 2px 9px; border-radius: 6px; font-weight: 600; font-size: 12.5px; white-space: nowrap; }
  .horas-baja { background: #dcfce7; color: #166534; }
  .horas-media { background: #fef3c7; color: #92400e; }
  .horas-alta { background: #fee2e2; color: #991b1b; }
  .vacio { background: #fff; border: 1px solid var(--borde); border-radius: 10px; padding: 16px; font-weight: 600; color: #166534; }
  @media print { header.encabezado { -webkit-print-color-adjust: exact; print-color-adjust: exact; } .tabla-wrap { max-height: none; overflow: visible; } .filtros { display: none; } }
</style>
'@

    $pre = '<header class="encabezado"><h1>Mails sin responder</h1>' +
        "<p class=`"fecha`">Día analizado: $stamp &middot; Generado: $((Get-Date).ToString('yyyy-MM-dd HH:mm'))</p></header>" +
        '<main><h2>Resumen por casilla</h2>' + (New-ResumenHtml $stats)
    if ($rows.Count -eq 0) { $pre += '<p class="vacio">No quedó ningún mail sin responder.</p>' }

    $documento = "<!DOCTYPE html><html lang=`"es`"><head>$head<title>Mails sin responder $stamp</title></head><body>" +
        $pre + (New-DetalleHtml $rows $stats) + '</main></body></html>'
    Set-Content -Path $htmlPath -Value $documento -Encoding UTF8

    return $csvPath, $htmlPath
}


function Main {
    $dayStart = $Fecha.Date
    $dayEnd = $dayStart.AddDays(1)
    Write-Host "Analizando mails del $($dayStart.ToString('yyyy-MM-dd'))..."

    try {
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace('MAPI')
    } catch {
        Write-Error ("No pude conectarme a Outlook. Verificá que el Outlook de escritorio esté instalado y abierto. " +
            "Detalle: $($_.Exception.Message)")
        exit 1
    }

    $me = Get-MyAddresses $namespace
    $toSent = { param($m, $f) Convert-Sent $m $f }
    $toReceived = { param($m, $f) Convert-Received $m $f $owner }

    # Tus Enviados: si respondiste "como" la casilla y Exchange guardó la copia solo en los tuyos.
    Write-Host 'Leyendo tus Elementos enviados...'
    $mySent = Read-Folder $namespace.GetDefaultFolder($olFolderSentMail) 'SentOn' $dayStart $dayEnd $IncluirSubcarpetas $toSent

    $allRows = @()
    $allStats = @()
    foreach ($name in $Casillas) {
        try {
            $mailbox = Open-Mailbox $namespace $name $me
        } catch {
            Write-Warning "Se saltea la casilla '$name': $($_.Exception.Message)"
            continue
        }
        $owner = $mailbox.Direcciones

        Write-Host "[$($mailbox.Nombre)] Leyendo Bandeja de entrada..."
        $received = Read-Folder $mailbox.Inbox 'ReceivedTime' $dayStart $dayEnd $IncluirSubcarpetas $toReceived

        $sent = New-Object System.Collections.Generic.List[object]
        foreach ($s in $mySent) { $sent.Add($s) }
        if ($mailbox.Enviados -and $mailbox.Nombre -ne 'Personal') {
            Write-Host "[$($mailbox.Nombre)] Leyendo Elementos enviados..."
            try {
                foreach ($s in (Read-Folder $mailbox.Enviados 'SentOn' $dayStart $dayEnd $IncluirSubcarpetas $toSent)) { $sent.Add($s) }
            } catch {
                Write-Warning "No pude leer los Enviados de '$($mailbox.Nombre)': $($_.Exception.Message)"
            }
        }

        $team = New-AddressSet
        $team.UnionWith($me)
        $team.UnionWith($owner)
        foreach ($address in $Equipo) { if ($address) { [void]$team.Add($address) } }

        $result = Find-Unanswered $received $sent $team $dayEnd $mailbox.Nombre
        $allRows += $result.Filas
        $allStats += $result.Estadisticas
    }

    if ($allStats.Count -eq 0) {
        Write-Error 'No se pudo leer ninguna casilla. Revisá la lista $Casillas en la configuración.'
        exit 1
    }

    $csvPath, $htmlPath = Write-Report $allRows $allStats $dayStart $CarpetaSalida

    Write-Host ''
    $allStats | Format-Table -AutoSize | Out-String | Write-Host
    Write-Host "CSV:  $csvPath"
    Write-Host "HTML: $htmlPath"

    if ($AbrirReporteAlTerminar) { Invoke-Item $htmlPath }
}


# Si el script se carga con ". .\reporte_no_respondidos.ps1" (para pruebas), no corre Main.
if ($MyInvocation.InvocationName -ne '.') { Main }
