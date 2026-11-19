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
  /* Tonos suaves para leer muchas horas: fondo cálido, texto no tan negro, poco brillo. */
  :root {
    --navy: #0C2A3A; --azul: #117ACA; --texto: #2d333b; --gris: #5f6b7a; --borde: #e4e0d8; --fondo: #f6f4ef;
    --superficie: #fffdf9; --superficie-2: #faf8f3; --cabecera: #eef2f5; --hover: #f1f5f9; --boton-borde: #cfc9bd;
    --rojo: #b42318; --verde-txt: #2f6b3f;
    --para-bg: #e3f3e8; --para-txt: #2f6b3f; --copia-bg: #e4edfb; --copia-txt: #2c4f8a; --otro-bg: #efeee9; --otro-txt: #55595f;
    --baja-bg: #e3f3e8; --baja-txt: #2f6b3f; --media-bg: #fbf0d4; --media-txt: #7a5a14; --alta-bg: #f9e1de; --alta-txt: #8f2a22;
  }
  html[data-tema="oscuro"] {
    --navy: #0b2331; --azul: #5aa9e6; --texto: #d4dae1; --gris: #98a2ae; --borde: #2f3742; --fondo: #14181e;
    --superficie: #1b2028; --superficie-2: #1f252e; --cabecera: #232b36; --hover: #263040; --boton-borde: #3d4655;
    --rojo: #ff9a90; --verde-txt: #9fdcb0;
    --para-bg: #1d3a28; --para-txt: #9fdcb0; --copia-bg: #1d3046; --copia-txt: #a9cdf2; --otro-bg: #272d36; --otro-txt: #c3cad3;
    --baja-bg: #1d3a28; --baja-txt: #9fdcb0; --media-bg: #3a2f17; --media-txt: #f0cf7a; --alta-bg: #3b2020; --alta-txt: #f2b8b2;
  }
  @media (prefers-color-scheme: dark) {
    html:not([data-tema="claro"]) {
      --navy: #0b2331; --azul: #5aa9e6; --texto: #d4dae1; --gris: #98a2ae; --borde: #2f3742; --fondo: #14181e;
      --superficie: #1b2028; --superficie-2: #1f252e; --cabecera: #232b36; --hover: #263040; --boton-borde: #3d4655;
      --rojo: #ff9a90; --verde-txt: #9fdcb0;
      --para-bg: #1d3a28; --para-txt: #9fdcb0; --copia-bg: #1d3046; --copia-txt: #a9cdf2; --otro-bg: #272d36; --otro-txt: #c3cad3;
      --baja-bg: #1d3a28; --baja-txt: #9fdcb0; --media-bg: #3a2f17; --media-txt: #f0cf7a; --alta-bg: #3b2020; --alta-txt: #f2b8b2;
    }
  }
  * { box-sizing: border-box; }
  body { margin: 0; font-family: "Segoe UI", Roboto, Arial, sans-serif; font-size: 17px; line-height: 1.6; color: var(--texto); background: var(--fondo); }
  header.encabezado { background: var(--navy); color: #f4f7fa; padding: 26px 36px; display: flex; justify-content: space-between; align-items: center; gap: 16px; flex-wrap: wrap; }
  header.encabezado h1 { margin: 0; font-size: 28px; font-weight: 600; }
  header.encabezado .fecha { margin: 6px 0 0; color: #d6e4ef; font-size: 16px; }
  .boton-tema { padding: 9px 16px; border: 1px solid #c9dbea; background: transparent; color: #f4f7fa; border-radius: 999px; cursor: pointer; font: inherit; font-size: 15px; }
  .boton-tema:hover { background: rgba(255,255,255,.12); }
  main { max-width: 1400px; margin: 0 auto; padding: 26px 36px 48px; }
  h2 { font-size: 21px; font-weight: 600; margin: 30px 0 14px; }
  .tarjetas { display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 18px; }
  .tarjeta { background: var(--superficie); border: 1px solid var(--borde); border-radius: 12px; padding: 18px 20px; }
  .tarjeta.total { border-color: var(--azul); }
  .tarjeta-titulo { font-weight: 600; margin-bottom: 12px; font-size: 17px; word-break: break-word; }
  .metricas { display: flex; gap: 24px; flex-wrap: wrap; }
  .metricas > div { display: flex; flex-direction: column; }
  .num { font-size: 30px; font-weight: 700; line-height: 1.15; }
  .lbl { font-size: 14px; color: var(--gris); letter-spacing: .02em; }
  .destacado .num { color: var(--rojo); }
  .nota { margin-top: 12px; font-size: 15px; color: var(--gris); }
  .filtros { background: var(--superficie); border: 1px solid var(--borde); border-radius: 12px; padding: 12px 18px; }
  .grupo { display: flex; flex-wrap: wrap; align-items: center; gap: 10px; margin: 10px 0; }
  .etiqueta { width: 84px; font-size: 14px; font-weight: 600; color: var(--gris); }
  .filtros button { padding: 8px 16px; border: 1px solid var(--boton-borde); background: var(--superficie); color: var(--texto); border-radius: 999px; cursor: pointer; font: inherit; font-size: 15px; }
  .filtros button:hover { border-color: var(--azul); }
  .filtros button.activo { background: var(--azul); border-color: var(--azul); color: #fff; }
  .contador { margin: 14px 2px 6px; color: var(--gris); font-size: 15px; }
  .leyenda { margin: 6px 2px 12px; color: var(--gris); font-size: 15px; }
  .leyenda .horas { margin-right: 4px; }
  .tabla-wrap { overflow: auto; max-height: 66vh; border: 1px solid var(--borde); border-radius: 12px; background: var(--superficie); }
  table { border-collapse: collapse; width: 100%; font-size: 16px; }
  thead th { position: sticky; top: 0; background: var(--cabecera); text-align: left; font-weight: 600; font-size: 15px; padding: 13px 14px; border-bottom: 1px solid var(--borde); white-space: nowrap; z-index: 1; }
  td { padding: 13px 14px; border-bottom: 1px solid var(--borde); vertical-align: top; line-height: 1.5; }
  tr.par td { background: var(--superficie-2); }
  tbody tr:hover td { background: var(--hover); }
  .nowrap { white-space: nowrap; }
  .email { color: var(--gris); word-break: break-all; font-size: 15px; }
  .asunto { min-width: 280px; }
  .num-col { text-align: right; white-space: nowrap; }
  .badge { display: inline-block; padding: 3px 11px; border-radius: 999px; font-size: 14px; font-weight: 600; white-space: nowrap; }
  .badge-para { background: var(--para-bg); color: var(--para-txt); }
  .badge-copia { background: var(--copia-bg); color: var(--copia-txt); }
  .badge-otro { background: var(--otro-bg); color: var(--otro-txt); }
  .horas { display: inline-block; padding: 3px 10px; border-radius: 6px; font-weight: 600; font-size: 15px; white-space: nowrap; }
  .horas-baja { background: var(--baja-bg); color: var(--baja-txt); }
  .horas-media { background: var(--media-bg); color: var(--media-txt); }
  .horas-alta { background: var(--alta-bg); color: var(--alta-txt); }
  .vacio { background: var(--superficie); border: 1px solid var(--borde); border-radius: 12px; padding: 18px; font-weight: 600; font-size: 17px; color: var(--verde-txt); }
  @media print {
    header.encabezado { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    .tabla-wrap { max-height: none; overflow: visible; }
    .filtros, .boton-tema { display: none; }
  }
</style>
'@

    $pre = '<header class="encabezado"><div><h1>Mails sin responder</h1>' +
        "<p class=`"fecha`">Día analizado: $stamp &middot; Generado: $((Get-Date).ToString('yyyy-MM-dd HH:mm'))</p></div>" +
        '<button id="tema" type="button" class="boton-tema">Modo oscuro</button></header>' +
        '<main><h2>Resumen por casilla</h2>' + (New-ResumenHtml $stats)
    if ($rows.Count -eq 0) { $pre += '<p class="vacio">No quedó ningún mail sin responder.</p>' }

    $jsTema = @'
<script>
(function () {
  var html = document.documentElement;
  var btn = document.getElementById('tema');
  function esOscuro() {
    if (html.dataset.tema) return html.dataset.tema === 'oscuro';
    return window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches;
  }
  function actualizarBoton() { btn.textContent = esOscuro() ? 'Modo claro' : 'Modo oscuro'; }
  try { var guardado = localStorage.getItem('tema-reporte'); if (guardado) html.dataset.tema = guardado; } catch (e) {}
  actualizarBoton();
  btn.addEventListener('click', function () {
    var nuevo = esOscuro() ? 'claro' : 'oscuro';
    html.dataset.tema = nuevo;
    try { localStorage.setItem('tema-reporte', nuevo); } catch (e) {}
    actualizarBoton();
  });
})();
</script>
'@

    $documento = "<!DOCTYPE html><html lang=`"es`"><head>$head<title>Mails sin responder $stamp</title></head><body>" +
        $pre + (New-DetalleHtml $rows $stats) + '</main>' + $jsTema + '</body></html>'
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
