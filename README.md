<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, viewport-fit=cover">
  <title>Control de Asistencia QR</title>
  <!-- Tailwind CSS -->
  <script src="https://cdn.tailwindcss.com"></script>
  <!-- Lector QR Optimizado -->
  <script src="https://unpkg.com/html5-qrcode" type="text/javascript"></script>
  <!-- Exportador Excel -->
  <script src="https://cdn.sheetjs.com/xlsx-latest/package/dist/xlsx.full.min.js"></script>
  <!-- Iconos FontAwesome -->
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
  <style>
    /* Estilos para vista de cámara en móvil */
    #reader video {
      object-fit: cover !important;
      border-radius: 1rem;
    }
    #reader {
      border: none !important;
    }
    .pb-safe {
      padding-bottom: env(safe-area-inset-bottom, 1rem);
    }
  </style>
</head>
<body class="bg-slate-50 text-slate-800 antialiased min-h-screen flex flex-col justify-between select-none">

  <!-- ENCABEZADO MÓVIL COMPACTO -->
  <header class="bg-indigo-700 text-white px-4 py-3 sticky top-0 z-40 shadow-sm flex items-center justify-between">
    <div class="flex items-center gap-2">
      <div class="bg-indigo-600 p-2 rounded-lg">
        <i class="fa-solid fa-qrcode text-lg"></i>
      </div>
      <div>
        <h1 class="font-bold text-base leading-tight">Pase de Lista QR</h1>
        <p id="txt-subtitulo-fecha" class="text-[11px] text-indigo-200">Cargando fecha...</p>
      </div>
    </div>
    <div id="badge-hoy" class="bg-indigo-800 text-indigo-100 text-xs px-2.5 py-1 rounded-full font-semibold">
      0 hoy
    </div>
  </header>

  <!-- CONTENEDOR PRINCIPAL CON ESPACIO PARA BARRA INFERIOR -->
  <main class="flex-1 max-w-lg w-full mx-auto p-4 pb-28">

    <!-- 1. PESTAÑA: ESCÁNER (VISTA PRINCIPAL MÓVIL) -->
    <section id="sec-escaner" class="space-y-4">
      <!-- Visor de Cámara -->
      <div class="bg-white p-3 rounded-2xl shadow-sm border border-slate-200 text-center">
        <div class="relative overflow-hidden rounded-xl bg-slate-950 aspect-[4/3] flex items-center justify-center">
          <div id="reader" class="w-full h-full"></div>
          <div id="placeholder-camara" class="absolute inset-0 flex flex-col items-center justify-center text-slate-400 p-4">
            <i class="fa-solid fa-camera text-4xl mb-2 text-slate-500"></i>
            <p class="text-xs">Cámara apagada para ahorrar batería</p>
          </div>
        </div>

        <!-- Controles de Cámara -->
        <div class="mt-3 flex gap-2">
          <button id="btn-iniciar-cam" onclick="toggleCamara()" class="flex-1 bg-indigo-600 active:bg-indigo-800 text-white font-semibold py-3 px-4 rounded-xl text-sm flex items-center justify-center gap-2 shadow-sm transition">
            <i class="fa-solid fa-play" id="icon-cam"></i>
            <span id="txt-btn-cam">Encender Cámara</span>
          </button>
        </div>

        <!-- Alerta de Escaneo Flotante -->
        <div id="notificacion" class="mt-3 hidden p-3 rounded-xl text-xs font-semibold border flex items-center gap-2 text-left"></div>
      </div>

      <!-- Resumen del Día Desplegable -->
      <div class="bg-white rounded-2xl shadow-sm border border-slate-200 p-4">
        <div class="flex justify-between items-center mb-2">
          <h2 class="text-xs font-bold uppercase tracking-wider text-slate-500">Registrados Hoy</h2>
          <span id="contador-hoy-detalle" class="text-xs font-bold text-indigo-600">0 alumnos</span>
        </div>

        <div class="overflow-y-auto max-h-48 divide-y divide-slate-100 text-sm" id="lista-hoy">
          <p class="text-center py-6 text-slate-400 text-xs">Sin asistencias escaneadas hoy.</p>
        </div>
      </div>
    </section>

    <!-- 2. PESTAÑA: PADRÓN DE ALUMNOS -->
    <section id="sec-alumnos" class="hidden space-y-4">
      <div class="bg-white p-4 rounded-2xl shadow-sm border border-slate-200">
        <h2 class="font-bold text-slate-800 text-sm mb-3">Registrar Nuevo Alumno</h2>
        <form id="form-alumno" onsubmit="guardarAlumno(event)" class="space-y-3">
          <div>
            <label class="text-[11px] font-bold text-slate-500 uppercase">Matrícula (Valor del QR)</label>
            <input type="text" id="al-matricula" required placeholder="Ej. 2026101" class="w-full text-base border border-slate-300 rounded-xl p-2.5 outline-none focus:border-indigo-600 focus:ring-1 focus:ring-indigo-600">
          </div>
          <div>
            <label class="text-[11px] font-bold text-slate-500 uppercase">Nombre Completo</label>
            <input type="text" id="al-nombre" required placeholder="Ej. Juan Pérez Martínez" class="w-full text-base border border-slate-300 rounded-xl p-2.5 outline-none focus:border-indigo-600 focus:ring-1 focus:ring-indigo-600">
          </div>
          <div>
            <label class="text-[11px] font-bold text-slate-500 uppercase">Grupo / Semestre</label>
            <input type="text" id="al-grupo" required placeholder="Ej. 2A" class="w-full text-base border border-slate-300 rounded-xl p-2.5 outline-none focus:border-indigo-600 focus:ring-1 focus:ring-indigo-600">
          </div>
          <button type="submit" class="w-full bg-indigo-600 active:bg-indigo-700 text-white font-semibold py-2.5 rounded-xl text-sm">
            Guardar en Padrón
          </button>
        </form>

        <!-- Acordeón Carga Masiva -->
        <details class="mt-4 border-t pt-3 border-slate-100 text-xs">
          <summary class="font-bold text-indigo-600 cursor-pointer py-1">Cargar Lista Completa (CSV / Texto)</summary>
          <div class="pt-2 space-y-2">
            <p class="text-slate-500 text-[11px]">Formato: <code>Matrícula,Nombre,Grupo</code> (uno por fila)</p>
            <textarea id="csv-masivo" rows="3" placeholder="202601,Ana López,2A&#10;202602,Pedro Gómez,2A" class="w-full text-xs font-mono border rounded-lg p-2"></textarea>
            <button onclick="cargarMasivo()" class="w-full bg-slate-800 text-white py-2 rounded-lg font-medium">Importar Lista</button>
          </div>
        </details>
      </div>

      <!-- Directorio -->
      <div class="bg-white p-4 rounded-2xl shadow-sm border border-slate-200">
        <div class="flex justify-between items-center mb-3">
          <h2 class="font-bold text-slate-800 text-sm">Alumnos en Sistema (<span id="total-alumnos-num">0</span>)</h2>
        </div>
        <input type="text" id="buscar-alumno" oninput="filtrarAlumnosDirectorio()" placeholder="Buscar alumno..." class="w-full text-xs border rounded-lg p-2 mb-3 bg-slate-50">
        <div class="overflow-y-auto max-h-60 divide-y divide-slate-100 text-xs" id="lista-directorio"></div>
      </div>
    </section>

    <!-- 3. PESTAÑA: CONCENTRADO Y REPORTES -->
    <section id="sec-reportes" class="hidden space-y-4">
      <div class="bg-white p-4 rounded-2xl shadow-sm border border-slate-200">
        <h2 class="font-bold text-slate-800 text-sm mb-3">Rango de Evaluación</h2>
        <div class="grid grid-cols-2 gap-2 mb-2">
          <div>
            <label class="text-[11px] font-semibold text-slate-500">Desde</label>
            <input type="date" id="rep-inicio" class="w-full text-xs border rounded-lg p-2 focus:ring-1 focus:ring-indigo-600">
          </div>
          <div>
            <label class="text-[11px] font-semibold text-slate-500">Hasta</label>
            <input type="date" id="rep-fin" class="w-full text-xs border rounded-lg p-2 focus:ring-1 focus:ring-indigo-600">
          </div>
        </div>

        <div class="mb-3">
          <label class="text-[11px] font-semibold text-slate-500">Grupo</label>
          <select id="rep-grupo" class="w-full text-xs border rounded-lg p-2 bg-white">
            <option value="">Todos los grupos</option>
          </select>
        </div>

        <div class="flex gap-2">
          <button onclick="generarReporte()" class="flex-1 bg-indigo-600 active:bg-indigo-700 text-white font-semibold py-2.5 rounded-xl text-xs">
            <i class="fa-solid fa-calculator mr-1"></i> Calcular
          </button>
          <button onclick="exportarExcel()" id="btn-excel" disabled class="flex-1 bg-emerald-600 disabled:bg-slate-300 text-white font-semibold py-2.5 rounded-xl text-xs">
            <i class="fa-solid fa-file-excel mr-1"></i> Excel
          </button>
        </div>
      </div>

      <!-- Tarjetas de Resultados -->
      <div class="bg-white p-4 rounded-2xl shadow-sm border border-slate-200">
        <div class="flex justify-between items-center mb-2">
          <h3 class="font-bold text-slate-800 text-xs uppercase tracking-wider">Resultados</h3>
          <span id="txt-dias-laborados" class="text-[11px] font-semibold text-indigo-600"></span>
        </div>
        <div class="space-y-2 overflow-y-auto max-h-72" id="contenedor-tarjetas-reporte">
          <p class="text-center py-6 text-slate-400 text-xs">Selecciona fechas y presiona Calcular.</p>
        </div>
      </div>
    </section>

    <!-- 4. PESTAÑA: RESPALDO Y COMPARTIR CON OTROS DISPOSITIVOS -->
    <section id="sec-sincronizar" class="hidden space-y-4">
      <div class="bg-white p-4 rounded-2xl shadow-sm border border-slate-200 space-y-3">
        <div class="flex items-center gap-3">
          <div class="bg-emerald-100 text-emerald-700 p-3 rounded-xl text-lg">
            <i class="fa-solid fa-cloud-arrow-up"></i>
          </div>
          <div>
            <h2 class="font-bold text-slate-800 text-sm">Enviar a Otro Dispositivo</h2>
            <p class="text-xs text-slate-500">Pasa tus listas y asistencias a tu PC o a otro celular.</p>
          </div>
        </div>

        <p class="text-xs text-slate-600 bg-slate-50 p-3 rounded-xl border border-slate-100 leading-relaxed">
          Genera un archivo con toda tu información. Puedes enviarlo directamente por <b>WhatsApp, Telegram, correo o Drive</b>.
        </p>

        <div class="flex flex-col gap-2">
          <button onclick="compartirRespaldoNativo()" class="w-full bg-emerald-600 active:bg-emerald-700 text-white font-semibold py-3 rounded-xl text-xs flex items-center justify-center gap-2">
            <i class="fa-solid fa-share-nodes text-sm"></i> Compartir Respaldo (WhatsApp / Drive)
          </button>
          <button onclick="descargarRespaldoArchivo()" class="w-full bg-slate-100 active:bg-slate-200 text-slate-700 font-semibold py-2.5 rounded-xl text-xs">
            <i class="fa-solid fa-download mr-1"></i> Descargar archivo (.json)
          </button>
        </div>
      </div>

      <div class="bg-white p-4 rounded-2xl shadow-sm border border-slate-200 space-y-3">
        <div class="flex items-center gap-3">
          <div class="bg-indigo-100 text-indigo-700 p-3 rounded-xl text-lg">
            <i class="fa-solid fa-cloud-arrow-down"></i>
          </div>
          <div>
            <h2 class="font-bold text-slate-800 text-sm">Restaurar en Este Dispositivo</h2>
            <p class="text-xs text-slate-500">Carga un respaldo recibido desde otro celular o PC.</p>
          </div>
        </div>

        <input type="file" id="input-archivo-backup" accept=".json" class="hidden" onchange="procesarArchivoRestaurar(event)">
        <button onclick="document.getElementById('input-archivo-backup').click()" class="w-full bg-indigo-600 active:bg-indigo-700 text-white font-semibold py-3 rounded-xl text-xs flex items-center justify-center gap-2">
          <i class="fa-solid fa-folder-open text-sm"></i> Seleccionar Archivo de Respaldo (.json)
        </button>
      </div>

      <!-- Limpieza de Seguridad -->
      <div class="p-4 rounded-2xl border border-rose-200 bg-rose-50">
        <h3 class="font-bold text-rose-800 text-xs mb-1">Borrar Base de Datos Local</h3>
        <p class="text-[11px] text-rose-600 mb-2">Elimina el registro de asistencias y alumnos en este celular.</p>
        <button onclick="borrarTodo()" class="bg-rose-600 active:bg-rose-700 text-white text-xs font-semibold py-2 px-3 rounded-lg">
          Vaciar Todo
        </button>
      </div>
    </section>

  </main>

  <!-- BARRA DE NAVEGACIÓN INFERIOR ESTILO APP (THUMB-FRIENDLY) -->
  <nav class="fixed bottom-0 left-0 right-0 bg-white border-t border-slate-200 z-50 px-2 py-1 shadow-lg pb-safe">
    <div class="max-w-lg mx-auto grid grid-cols-4 gap-1 text-center">
      <button onclick="cambiarPestana('escaner')" id="nav-escaner" class="py-2 flex flex-col items-center text-indigo-600">
        <i class="fa-solid fa-qrcode text-lg"></i>
        <span class="text-[10px] font-bold mt-1">Escanear</span>
      </button>
      <button onclick="cambiarPestana('alumnos')" id="nav-alumnos" class="py-2 flex flex-col items-center text-slate-400">
        <i class="fa-solid fa-users text-lg"></i>
        <span class="text-[10px] font-medium mt-1">Alumnos</span>
      </button>
      <button onclick="cambiarPestana('reportes')" id="nav-reportes" class="py-2 flex flex-col items-center text-slate-400">
        <i class="fa-solid fa-chart-pie text-lg"></i>
        <span class="text-[10px] font-medium mt-1">Reportes</span>
      </button>
      <button onclick="cambiarPestana('sincronizar')" id="nav-sincronizar" class="py-2 flex flex-col items-center text-slate-400">
        <i class="fa-solid fa-arrows-rotate text-lg"></i>
        <span class="text-[10px] font-medium mt-1">Sincronizar</span>
      </button>
    </div>
  </nav>

  <script>
    // --- ESTADO LOCAL ---
    let alumnos = JSON.parse(localStorage.getItem('qr_asistencia_alumnos')) || [];
    let registros = JSON.parse(localStorage.getItem('qr_asistencia_registros')) || [];
    let html5QrCode = null;
    let camaraEncendida = false;
    let datosUltimoReporte = [];

    // --- RESPUESTA HÁPTICA Y SONIDO EN MÓVIL ---
    function feedbackExito() {
      // Vibración de confirmación en móviles Android
      if (navigator.vibrate) navigator.vibrate([80, 50, 80]);
      reproducirAudio(880, 0.12);
    }

    function feedbackError() {
      if (navigator.vibrate) navigator.vibrate(250);
      reproducirAudio(280, 0.25);
    }

    function reproducirAudio(freq, duracion) {
      try {
        const audioCtx = new (window.AudioContext || window.webkitAudioContext)();
        const osc = audioCtx.createOscillator();
        const gain = audioCtx.createGain();
        osc.connect(gain);
        gain.connect(audioCtx.destination);
        osc.frequency.value = freq;
        gain.gain.setValueAtTime(0.15, audioCtx.currentTime);
        gain.gain.exponentialRampToValueAtTime(0.01, audioCtx.currentTime + duracion);
        osc.start();
        osc.stop(audioCtx.currentTime + duracion);
      } catch(e) {}
    }

    // --- NAVEGACIÓN TIPO APP ---
    function cambiarPestana(p) {
      ['escaner', 'alumnos', 'reportes', 'sincronizar'].forEach(sec => {
        document.getElementById(`sec-${sec}`).classList.add('hidden');
        const navBtn = document.getElementById(`nav-${sec}`);
        navBtn.className = "py-2 flex flex-col items-center text-slate-400";
        navBtn.querySelector('span').classList.replace('font-bold', 'font-medium');
      });

      document.getElementById(`sec-${p}`).classList.remove('hidden');
      const activo = document.getElementById(`nav-${p}`);
      activo.className = "py-2 flex flex-col items-center text-indigo-600";
      activo.querySelector('span').classList.replace('font-medium', 'font-bold');

      // Si sales del escáner, apagar cámara para ahorrar batería
      if (p !== 'escaner' && camaraEncendida) {
        toggleCamara();
      }
      if (p === 'reportes') poblarGrupos();
      if (p === 'alumnos') renderDirectorio();
    }

    // --- CÁMARA TRASERA EN MÓVIL ---
    function toggleCamara() {
      if (camaraEncendida) {
        apagarCamara();
      } else {
        encenderCamara();
      }
    }

    function encenderCamara() {
      if (!html5QrCode) {
        html5QrCode = new Html5Qrcode("reader");
      }
      document.getElementById('placeholder-camara').classList.add('hidden');

      // Configuración adaptada a teléfonos (prioriza cámara trasera)
      const config = {
        fps: 15,
        qrbox: (w, h) => ({ width: Math.floor(w * 0.7), height: Math.floor(w * 0.7) }),
        aspectRatio: 1.333
      };

      html5QrCode.start(
        { facingMode: "environment" },
        config,
        (decodedText) => procesarCodigo(decodedText.trim()),
        () => {}
      ).then(() => {
        camaraEncendida = true;
        document.getElementById('txt-btn-cam').innerText = "Pausar Cámara";
        document.getElementById('icon-cam').className = "fa-solid fa-pause";
        document.getElementById('btn-iniciar-cam').className = "flex-1 bg-amber-600 text-white font-semibold py-3 px-4 rounded-xl text-sm flex items-center justify-center gap-2";
      }).catch(err => {
        mostrarAviso("⚠️ No se pudo acceder a la cámara. Verifica los permisos de tu navegador.", "error");
        document.getElementById('placeholder-camara').classList.remove('hidden');
      });
    }

    function apagarCamara() {
      if (html5QrCode && camaraEncendida) {
        html5QrCode.stop().then(() => {
          camaraEncendida = false;
          document.getElementById('txt-btn-cam').innerText = "Encender Cámara";
          document.getElementById('icon-cam').className = "fa-solid fa-play";
          document.getElementById('btn-iniciar-cam').className = "flex-1 bg-indigo-600 text-white font-semibold py-3 px-4 rounded-xl text-sm flex items-center justify-center gap-2";
          document.getElementById('placeholder-camara').classList.remove('hidden');
        });
      }
    }

    // --- PROCESAMIENTO DE ASISTENCIA ---
    let ultimoEscaneo = 0;
    let ultimoCodigo = "";

    function procesarCodigo(codigo) {
      const ahora = Date.now();
      // Bloquear lecturas repetidas en menos de 3.5 segundos
      if (codigo === ultimoCodigo && (ahora - ultimoEscaneo) < 3500) return;
      ultimoCodigo = codigo;
      ultimoEscaneo = ahora;

      const hoyStr = new Date().toISOString().split('T')[0];
      const horaStr = new Date().toTimeString().split(' ')[0].substring(0, 5);

      // Comprobar si ya pasó lista hoy
      const yaRegistro = registros.some(r => r.matricula.toLowerCase() === codigo.toLowerCase() && r.fecha === hoyStr);
      if (yaRegistro) {
        feedbackError();
        mostrarAviso(`⚠️ ${codigo} YA REGISTRÓ ASISTENCIA HOY`, "warning");
        return;
      }

      // Buscar datos en el padrón
      const alumno = alumnos.find(a => a.matricula.toLowerCase() === codigo.toLowerCase());
      const nombre = alumno ? alumno.nombre : "Alumno Sin Registrar";
      const grupo = alumno ? alumno.grupo : "Sin grupo";

      const dias = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'];
      const diaSemana = dias[new Date().getDay()];

      registros.push({
        matricula: codigo,
        nombre: nombre,
        grupo: grupo,
        fecha: hoyStr,
        dia: diaSemana,
        hora: horaStr,
        ts: ahora
      });

      localStorage.setItem('qr_asistencia_registros', JSON.stringify(registros));
      feedbackExito();
      mostrarAviso(`✅ ${nombre} (${horaStr})`, "success");
      actualizarListaHoy();
    }

    function mostrarAviso(msj, tipo) {
      const el = document.getElementById('notificacion');
      el.classList.remove('hidden', 'bg-emerald-50', 'border-emerald-200', 'text-emerald-800', 
                            'bg-rose-50', 'border-rose-200', 'text-rose-800',
                            'bg-amber-50', 'border-amber-200', 'text-amber-800');

      if (tipo === 'success') el.classList.add('bg-emerald-50', 'border-emerald-200', 'text-emerald-800');
      else if (tipo === 'error') el.classList.add('bg-rose-50', 'border-rose-200', 'text-rose-800');
      else el.classList.add('bg-amber-50', 'border-amber-200', 'text-amber-800');

      el.innerHTML = msj;
    }

    function actualizarListaHoy() {
      const hoyStr = new Date().toISOString().split('T')[0];
      const hoyList = registros.filter(r => r.fecha === hoyStr).reverse();
      
      document.getElementById('badge-hoy').innerText = `${hoyList.length} hoy`;
      document.getElementById('contador-hoy-detalle').innerText = `${hoyList.length} asistencias`;

      const cont = document.getElementById('lista-hoy');
      if (hoyList.length === 0) {
        cont.innerHTML = '<p class="text-center py-6 text-slate-400 text-xs">Sin asistencias escaneadas hoy.</p>';
        return;
      }

      cont.innerHTML = hoyList.map(r => `
        <div class="py-2 flex justify-between items-center">
          <div>
            <p class="font-bold text-slate-800 text-xs">${r.nombre}</p>
            <p class="text-[10px] text-slate-400 font-mono">${r.matricula} · ${r.grupo}</p>
          </div>
          <span class="bg-indigo-50 text-indigo-700 px-2 py-1 rounded-md text-[11px] font-bold">${r.hora}</span>
        </div>
      `).join('');
    }

    // --- PADRÓN DE ALUMNOS ---
    function guardarAlumno(e) {
      e.preventDefault();
      const mat = document.getElementById('al-matricula').value.trim();
      const nom = document.getElementById('al-nombre').value.trim();
      const grp = document.getElementById('al-grupo').value.trim();

      const idx = alumnos.findIndex(a => a.matricula.toLowerCase() === mat.toLowerCase());
      if (idx >= 0) alumnos[idx] = { matricula: mat, nombre: nom, grupo: grp };
      else alumnos.push({ matricula: mat, nombre: nom, grupo: grp });

      localStorage.setItem('qr_asistencia_alumnos', JSON.stringify(alumnos));
      document.getElementById('form-alumno').reset();
      renderDirectorio();
      alert("Alumno guardado con éxito.");
    }

    function cargarMasivo() {
      const txt = document.getElementById('csv-masivo').value.trim();
      if (!txt) return;
      let count = 0;
      txt.split('\n').forEach(linea => {
        const p = linea.split(',');
        if (p.length >= 2) {
          const mat = p[0].trim();
          const nom = p[1].trim();
          const grp = (p[2] || 'A').trim();
          if (mat && nom) {
            const idx = alumnos.findIndex(a => a.matricula.toLowerCase() === mat.toLowerCase());
            if (idx >= 0) alumnos[idx] = { matricula: mat, nombre: nom, grupo: grp };
            else alumnos.push({ matricula: mat, nombre: nom, grupo: grp });
            count++;
          }
        }
      });
      localStorage.setItem('qr_asistencia_alumnos', JSON.stringify(alumnos));
      document.getElementById('csv-masivo').value = '';
      renderDirectorio();
      alert(`Se importaron ${count} alumnos.`);
    }

    function renderDirectorio() {
      document.getElementById('total-alumnos-num').innerText = alumnos.length;
      filtrarAlumnosDirectorio();
    }

    function filtrarAlumnosDirectorio() {
      const q = (document.getElementById('buscar-alumno').value || '').toLowerCase();
      const lista = alumnos.filter(a => a.nombre.toLowerCase().includes(q) || a.matricula.toLowerCase().includes(q));
      const cont = document.getElementById('lista-directorio');

      if (lista.length === 0) {
        cont.innerHTML = '<p class="text-center py-4 text-slate-400">Sin coincidencias.</p>';
        return;
      }

      cont.innerHTML = lista.map(a => `
        <div class="py-2 flex justify-between items-center">
          <div>
            <p class="font-bold text-slate-800 text-xs">${a.nombre}</p>
            <p class="text-[10px] text-slate-400">${a.matricula} · Grupo ${a.grupo}</p>
          </div>
          <button onclick="eliminarAlumno('${a.matricula}')" class="text-rose-500 p-2"><i class="fa-solid fa-trash-can"></i></button>
        </div>
      `).join('');
    }

    function eliminarAlumno(mat) {
      if (confirm(`¿Eliminar matrícula ${mat}?`)) {
        alumnos = alumnos.filter(a => a.matricula !== mat);
        localStorage.setItem('qr_asistencia_alumnos', JSON.stringify(alumnos));
        renderDirectorio();
      }
    }

    function poblarGrupos() {
      const grps = [...new Set(alumnos.map(a => a.grupo))].filter(Boolean);
      document.getElementById('rep-grupo').innerHTML = '<option value="">Todos los grupos</option>' + 
        grps.map(g => `<option value="${g}">${g}</option>`).join('');
    }

    // --- REPORTE Y CONCENTRADO ---
    function generarReporte() {
      const ini = document.getElementById('rep-inicio').value;
      const fin = document.getElementById('rep-fin').value;
      const grp = document.getElementById('rep-grupo').value;

      if (!ini || !fin) return alert("Selecciona fecha inicial y final.");
      if (ini > fin) return alert("La fecha inicial no puede ser mayor.");

      // Días únicos con actividad en ese periodo
      const diasLaborados = [...new Set(registros.filter(r => r.fecha >= ini && r.fecha <= fin).map(r => r.fecha))];
      const totalDias = diasLaborados.length;
      document.getElementById('txt-dias-laborados').innerText = `${totalDias} Días Activos`;

      let lista = alumnos;
      if (grp) lista = lista.filter(a => a.grupo === grp);

      if (lista.length === 0) {
        document.getElementById('contenedor-tarjetas-reporte').innerHTML = '<p class="text-center py-6 text-slate-400 text-xs">No hay alumnos registrados en el grupo seleccionado.</p>';
        document.getElementById('btn-excel').disabled = true;
        return;
      }

      datosUltimoReporte = lista.map(al => {
        const asistencias = registros.filter(r => r.matricula.toLowerCase() === al.matricula.toLowerCase() && r.fecha >= ini && r.fecha <= fin).length;
        const inasistencias = Math.max(0, totalDias - asistencias);
        const pct = totalDias > 0 ? ((asistencias / totalDias) * 100).toFixed(1) : 0;

        return {
          matricula: al.matricula,
          nombre: al.nombre,
          grupo: al.grupo,
          asistencias: asistencias,
          inasistencias: inasistencias,
          totalDias: totalDias,
          porcentaje: `${pct}%`
        };
      });

      const cont = document.getElementById('contenedor-tarjetas-reporte');
      cont.innerHTML = datosUltimoReporte.map(d => {
        const pctNum = parseFloat(d.porcentaje);
        const color = pctNum >= 80 ? 'text-emerald-600 bg-emerald-50' : (pctNum >= 60 ? 'text-amber-600 bg-amber-50' : 'text-rose-600 bg-rose-50');

        return `
          <div class="p-2.5 rounded-xl border border-slate-100 bg-slate-50 flex items-center justify-between text-xs">
            <div>
              <p class="font-bold text-slate-800">${d.nombre}</p>
              <p class="text-[10px] text-slate-400">${d.matricula} · Grupo ${d.grupo}</p>
              <p class="text-[11px] mt-1">
                <span class="text-emerald-700 font-semibold">${d.asistencias} Asist.</span> · 
                <span class="text-rose-700 font-semibold">${d.inasistencias} Inasist.</span>
              </p>
            </div>
            <span class="font-bold px-2 py-1 rounded-lg text-xs ${color}">${d.porcentaje}</span>
          </div>
        `;
      }).join('');

      document.getElementById('btn-excel').disabled = false;
    }

    function exportarExcel() {
      if (datosUltimoReporte.length === 0) return;
      const datos = datosUltimoReporte.map(d => ({
        "Matrícula": d.matricula,
        "Nombre": d.nombre,
        "Grupo": d.grupo,
        "Total Días Lectivos": d.totalDias,
        "Asistencias": d.asistencias,
        "Inasistencias": d.inasistencias,
        "% Cumplimiento": d.porcentaje
      }));
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.json_to_sheet(datos);
      XLSX.utils.book_append_sheet(wb, ws, "Concentrado");
      XLSX.writeFile(wb, `Concentrado_Asistencias_${document.getElementById('rep-inicio').value}.xlsx`);
    }

    // --- COMPARTIR Y SINCRONIZAR MULTIDISPOSITIVO ---
    function obtenerPayloadRespaldo() {
      return JSON.stringify({
        version: "1.0",
        fechaExportacion: new Date().toISOString(),
        alumnos: alumnos,
        registros: registros
      }, null, 2);
    }

    async function compartirRespaldoNativo() {
      const jsonStr = obtenerPayloadRespaldo();
      const fecha = new Date().toISOString().split('T')[0];
      const nombreArchivo = `asistencia_respaldo_${fecha}.json`;
      const blob = new Blob([jsonStr], { type: 'application/json' });
      const archivo = new File([blob], nombreArchivo, { type: 'application/json' });

      // Si el navegador soporta compartir archivos directamente (Chrome móvil, Safari iOS)
      if (navigator.canShare && navigator.canShare({ files: [archivo] })) {
        try {
          await navigator.share({
            files: [archivo],
            title: 'Respaldo de Asistencia QR',
            text: 'Aquí está la base de datos de alumnos y asistencias.'
          });
        } catch (e) {
          console.log("Compartir cancelado o no soportado");
        }
      } else {
        // Fallback: descargar el archivo para enviarlo manualmente
        descargarRespaldoArchivo();
        alert("Tu navegador no soporta el menú directo de compartir. El archivo de respaldo se ha descargado para que lo envíes por WhatsApp o correo.");
      }
    }

    function descargarRespaldoArchivo() {
      const jsonStr = obtenerPayloadRespaldo();
      const fecha = new Date().toISOString().split('T')[0];
      const blob = new Blob([jsonStr], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `asistencia_respaldo_${fecha}.json`;
      a.click();
      URL.revokeObjectURL(url);
    }

    function procesarArchivoRestaurar(e) {
      const file = e.target.files[0];
      if (!file) return;

      const reader = new FileReader();
      reader.onload = function(evt) {
        try {
          const data = JSON.parse(evt.target.result);
          if (!data.alumnos || !data.registros) throw new Error("Formato inválido");

          if (confirm(`Se encontraron ${data.alumnos.length} alumnos y ${data.registros.length} asistencias. ¿Deseas reemplazar los datos actuales con este respaldo?`)) {
            alumnos = data.alumnos;
            registros = data.registros;
            localStorage.setItem('qr_asistencia_alumnos', JSON.stringify(alumnos));
            localStorage.setItem('qr_asistencia_registros', JSON.stringify(registros));
            actualizarListaHoy();
            renderDirectorio();
            alert("¡Datos restaurados y sincronizados correctamente!");
          }
        } catch (err) {
          alert("Error: El archivo seleccionado no es un respaldo válido.");
        }
      };
      reader.readAsText(file);
    }

    function borrarTodo() {
      if (confirm("⚠️ ¿Estás seguro de vaciar todos los alumnos y asistencias de este dispositivo? Esta acción no se puede deshacer.")) {
        localStorage.clear();
        alumnos = [];
        registros = [];
        actualizarListaHoy();
        renderDirectorio();
        alert("Base de datos local vaciada.");
      }
    }

    // Inicialización
    window.addEventListener('DOMContentLoaded', () => {
      const hoy = new Date();
      document.getElementById('txt-subtitulo-fecha').innerText = hoy.toLocaleDateString('es-MX', { weekday: 'short', day: 'numeric', month: 'short' });
      
      const hoyISO = hoy.toISOString().split('T')[0];
      document.getElementById('rep-fin').value = hoyISO;
      document.getElementById('rep-inicio').value = new Date(hoy.getFullYear(), hoy.getMonth(), 1).toISOString().split('T')[0];

      actualizarListaHoy();
    });
  </script>
</body>
</html>
