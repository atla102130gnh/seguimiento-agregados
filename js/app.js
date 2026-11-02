/*
  ============================================================================
  Seguimiento de Carta de Agregados — lógica de la página
  ----------------------------------------------------------------------------
  Módulos (en este orden):
    1. Utilidades (fechas en hora de México, texto, avisos)
    2. Conexión a Supabase y sesión
    3. Catálogos en memoria
    4. Componentes: lista seleccionable, calendario, autocompletado tipo Excel
    5. Formulario de agregado (alta y edición)
    6. Consulta (tabla, filtros, indicadores, exportar)
    7. Acciones por registro (cargar, editar, baja, historial)
    8. Catálogos (pantalla)
    9. Correos y respaldo (pantalla)
   10. Navegación y arranque
  ============================================================================
*/
(function () {
  'use strict';

  /* =========================================================================
     1. UTILIDADES
     ========================================================================= */
  var TZ = 'America/Mexico_City';
  var DIAS_LIMITE = 4;
  var ESTATUS = { PEND: 'PENDIENTE', CARG: 'CARGADO', ACL: 'ACLARACIÓN SERVICIO NO BRINDADO' };
  var OPS = { OK: 'COMPLETADO', TIEMPO: 'EN TIEMPO', CAMBIO: 'CAMBIO DE ESTATUS OPERACIONES' };
  var MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];

  function $(sel, root) { return (root || document).querySelector(sel); }
  function $$(sel, root) { return Array.prototype.slice.call((root || document).querySelectorAll(sel)); }
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }
  // Normaliza para comparar: sin acentos, mayúsculas, sin espacios dobles
  function norm(s) {
    return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[\s\u00A0]+/g, ' ').trim().toUpperCase();
  }
  function compact(s) { return norm(s).replace(/[^A-Z0-9]/g, ''); }
  function limpiar(s) { return String(s == null ? '' : s).replace(/[\s\u00A0]+/g, ' ').trim(); }
  function debounce(fn, ms) { var t; return function () { var a = arguments, me = this; clearTimeout(t); t = setTimeout(function () { fn.apply(me, a); }, ms); }; }
  function icon(id, cls) { return '<svg' + (cls ? ' class="' + cls + '"' : '') + ' aria-hidden="true"><use href="#' + id + '"/></svg>'; }

  // Fechas: siempre 'AAAA-MM-DD' (como las guarda la base), en hora de México
  function hoyISO() {
    return new Intl.DateTimeFormat('en-CA', { timeZone: TZ, year: 'numeric', month: '2-digit', day: '2-digit' }).format(new Date());
  }
  function anioActual() { return +hoyISO().slice(0, 4); }
  function isoADMY(iso) { if (!iso) return ''; var p = String(iso).slice(0, 10).split('-'); return p[2] + '/' + p[1] + '/' + p[0]; }
  function isoValida(y, m, d) {
    var dt = new Date(Date.UTC(y, m - 1, d));
    return dt.getUTCFullYear() === y && dt.getUTCMonth() === m - 1 && dt.getUTCDate() === d;
  }
  function pad2(n) { return (n < 10 ? '0' : '') + n; }
  function aISO(y, m, d) { return y + '-' + pad2(m) + '-' + pad2(d); }
  // Acepta 5/3/26, 05/03/2026, 5-3-2026, 05032026 → '2026-03-05'
  function dmyAISO(txt) {
    var s = String(txt || '').trim();
    if (!s) return null;
    var m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2}|\d{4})$/) || s.match(/^(\d{2})(\d{2})(\d{4})$/);
    if (!m) return 'invalid';
    var d = +m[1], mo = +m[2], y = +m[3];
    if (m[3].length === 2) y = 2000 + y;
    if (y < 2000 || y > 2100 || !isoValida(y, mo, d)) return 'invalid';
    return aISO(y, mo, d);
  }
  function diasEntre(isoA, isoB) {
    var a = isoA.split('-'), b = isoB.split('-');
    return Math.round((Date.UTC(+b[0], b[1] - 1, +b[2]) - Date.UTC(+a[0], a[1] - 1, +a[2])) / 86400000);
  }
  function fechaHora(ts) {
    if (!ts) return '';
    return new Date(ts).toLocaleString('es-MX', { timeZone: TZ, day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });
  }

  // Estatus operaciones igual que la fórmula del Excel / la vista de la base
  function estatusOps(fechaSolicitud, estatus, fechaCarga) {
    if (fechaCarga || estatus === ESTATUS.CARG) return { ops: OPS.OK, dias: null };
    if (!fechaSolicitud) return { ops: '', dias: null };
    var d = diasEntre(fechaSolicitud, hoyISO());
    return { ops: d <= DIAS_LIMITE ? OPS.TIEMPO : OPS.CAMBIO, dias: d };
  }
  function badgeOps(ops, dias) {
    var cls = ops === OPS.OK ? 'ok' : ops === OPS.TIEMPO ? 'warn' : ops === OPS.CAMBIO ? 'bad' : 'neutral';
    var txt = ops === OPS.OK ? 'Completado' : ops === OPS.TIEMPO ? 'En tiempo' : ops === OPS.CAMBIO ? 'Cambio de estatus' : '—';
    return '<span class="badge ' + cls + '" title="' + esc(ops) + '">' + txt + '</span>' +
      (dias != null && ops !== OPS.OK ? '<div class="days">' + dias + (dias === 1 ? ' día' : ' días') + '</div>' : '');
  }
  function badgeEstatus(est) {
    if (est === ESTATUS.CARG) return '<span class="badge ok plain">Cargado</span>';
    if (est === ESTATUS.PEND) return '<span class="badge neutral plain">Pendiente</span>';
    return '<span class="badge info plain" title="' + esc(est) + '">Aclaración</span>';
  }

  // Avisos (toast)
  function toast(tipo, titulo, detalle) {
    var box = $('#toasts');
    var el = document.createElement('div');
    el.className = 'toast ' + tipo;
    el.setAttribute('role', tipo === 'bad' ? 'alert' : 'status');
    el.innerHTML = '<span class="t-ic">' + icon(tipo === 'ok' ? 'i-check' : tipo === 'bad' ? 'i-alert' : 'i-info') + '</span>' +
      '<div><b>' + esc(titulo) + '</b>' + (detalle ? '<span>' + esc(detalle) + '</span>' : '') + '</div>';
    box.appendChild(el);
    setTimeout(function () { el.remove(); }, tipo === 'bad' ? 9000 : 4500);
  }

  // Traduce errores de la base a mensajes claros
  function mensajeError(err) {
    if (!err) return 'Error desconocido.';
    var m = (err.message || String(err));
    var code = err.code || '';
    if (/Failed to fetch|NetworkError|network/i.test(m)) return 'Sin conexión con la base de datos. Revisa tu internet e inténtalo de nuevo.';
    if (code === '23505') return 'Ese nombre ya existe en el catálogo.';
    if (/chk_fecha_carga_posterior/.test(m)) return 'La fecha de carga no puede ser anterior a la fecha de solicitud.';
    if (/chk_sin_blancos/.test(m)) return 'Faltan datos obligatorios (contrarrecibo, NUR, CPT o descripción).';
    if (/chk_baja_completa/.test(m)) return 'Para dar de baja hay que escribir el motivo.';
    if (code === '42501' || /permission denied|row-level security/i.test(m)) return 'Tu usuario no tiene permiso para esta acción.';
    if (/JWT|token/i.test(m)) return 'Tu sesión expiró. Vuelve a iniciar sesión.';
    if (/No se permite borrar/.test(m)) return 'Los registros no se borran: usa “Dar de baja”.';
    return m;
  }

  // Diálogo de confirmación que devuelve una promesa
  function confirmar(titulo, html, textoSi, peligro) {
    return new Promise(function (resolve) {
      var dlg = $('#dlg-confirm');
      $('#dlg-confirm-title').textContent = titulo;
      $('#dlg-confirm-body').innerHTML = html;
      var si = $('#dlg-confirm-yes'), no = $('#dlg-confirm-no');
      si.textContent = textoSi || 'Continuar';
      si.className = peligro ? 'btn-danger' : 'btn-primary';
      function fin(v) { si.onclick = no.onclick = null; dlg.onclose = null; if (dlg.open) dlg.close(); resolve(v); }
      si.onclick = function () { fin(true); };
      no.onclick = function () { fin(false); };
      dlg.onclose = function () { resolve(false); };
      dlg.showModal();
      si.focus();
    });
  }

  function cargando(btn, on, texto) {
    if (!btn) return;
    if (on) { btn.dataset.html = btn.innerHTML; btn.disabled = true; btn.innerHTML = '<span class="spinner"></span>' + esc(texto || 'Guardando…'); }
    else { btn.disabled = false; if (btn.dataset.html) btn.innerHTML = btn.dataset.html; }
  }

  /* =========================================================================
     2. CONEXIÓN Y SESIÓN
     ========================================================================= */
  var CFG = window.APP_CONFIG || {};
  var sb = null;
  var state = {
    user: null,
    cat: { clientes: [], proveedores: [], solicitantes: [], tipos_servicio: [] },
    byId: { clientes: {}, proveedores: {}, solicitantes: {}, tipos_servicio: {} },
    anios: [],
    iniciado: {},
    consultaSucia: true
  };
  var CATALOGOS = [
    { key: 'clientes', titulo: 'Clientes', singular: 'cliente' },
    { key: 'proveedores', titulo: 'Proveedores', singular: 'proveedor' },
    { key: 'solicitantes', titulo: 'Solicitantes', singular: 'solicitante' },
    { key: 'tipos_servicio', titulo: 'Tipos de servicio', singular: 'tipo de servicio' }
  ];

  function configurado() {
    return CFG.SUPABASE_URL && CFG.SUPABASE_ANON_KEY &&
      !/TU-PROYECTO|PEGA_AQUI/i.test(CFG.SUPABASE_URL + CFG.SUPABASE_ANON_KEY) && window.supabase;
  }

  function mostrar(cual) {
    $('#setup').hidden = cual !== 'setup';
    $('#login').hidden = cual !== 'login';
    $('#app').hidden = cual !== 'app';
  }

  function iniciarSesionUI() {
    $('#login-form').addEventListener('submit', async function (e) {
      e.preventDefault();
      var email = $('#login-email').value.trim(), pass = $('#login-pass').value;
      var err = $('#login-error');
      err.hidden = true;
      if (!email || !pass) { err.textContent = 'Escribe tu correo y contraseña.'; err.hidden = false; return; }
      var btn = $('#login-btn');
      cargando(btn, true, 'Entrando…');
      var r = await sb.auth.signInWithPassword({ email: email, password: pass });
      cargando(btn, false);
      if (r.error) {
        err.textContent = /Invalid login/i.test(r.error.message) ? 'Correo o contraseña incorrectos.' : mensajeError(r.error);
        err.hidden = false;
      }
    });
    $('#logout').addEventListener('click', async function () {
      await sb.auth.signOut();
    });
  }

  async function alEntrar(session) {
    state.user = session.user;
    var email = state.user.email || '';
    var av = $('#user-avatar');
    av.textContent = email.slice(0, 2).toUpperCase();
    av.title = email;
    mostrar('app');
    try {
      await cargarCatalogos();
    } catch (e) {
      toast('bad', 'No se pudieron cargar los catálogos', mensajeError(e));
    }
    enrutar();
    actualizarContadorNav();
  }

  /* =========================================================================
     3. CATÁLOGOS EN MEMORIA
     ========================================================================= */
  async function cargarCatalogos() {
    var res = await Promise.all(CATALOGOS.map(function (c) {
      return sb.from(c.key).select('id,nombre,activo').order('nombre');
    }).concat([sb.from('v_anios_surtido').select('anio')]));
    CATALOGOS.forEach(function (c, i) {
      if (res[i].error) throw res[i].error;
      state.cat[c.key] = res[i].data;
      state.byId[c.key] = {};
      res[i].data.forEach(function (x) { state.byId[c.key][x.id] = x; });
    });
    var an = res[CATALOGOS.length];
    if (an.error) throw an.error;
    state.anios = an.data.map(function (x) { return x.anio; });
    if (state.anios.indexOf(anioActual()) < 0) state.anios.unshift(anioActual());
    state.anios.sort(function (a, b) { return b - a; });
    document.dispatchEvent(new CustomEvent('catalogos'));
  }
  function nombreDe(cat, id) { var x = state.byId[cat][id]; return x ? x.nombre : ''; }

  async function altaCatalogo(cat, nombre) {
    nombre = limpiar(nombre).toUpperCase();
    if (!nombre) throw new Error('Escribe un nombre.');
    var existe = state.cat[cat].filter(function (x) { return norm(x.nombre) === norm(nombre); })[0];
    if (existe) {
      if (!existe.activo) {
        var r0 = await sb.from(cat).update({ activo: true }).eq('id', existe.id).select().single();
        if (r0.error) throw r0.error;
        await cargarCatalogos();
        toast('ok', 'Reactivado en el catálogo', existe.nombre);
        return state.byId[cat][existe.id];
      }
      return existe;
    }
    var r = await sb.from(cat).insert({ nombre: nombre }).select().single();
    if (r.error) throw r.error;
    await cargarCatalogos();
    toast('ok', 'Dado de alta', nombre);
    return r.data;
  }

  // Busca nombres parecidos para evitar duplicados (STAR MEDICA vs STAR MÉDICA, etc.)
  function lev(a, b) {
    if (a === b) return 0;
    var m = a.length, n = b.length, prev = [], cur, i, j;
    for (j = 0; j <= n; j++) prev[j] = j;
    for (i = 1; i <= m; i++) {
      cur = [i];
      for (j = 1; j <= n; j++) cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
      prev = cur;
    }
    return prev[n];
  }
  var VACIAS = ['SA', 'DE', 'CV', 'DEL', 'LA', 'EL', 'LOS', 'LAS', 'Y', 'SC', 'SAS', 'RL', 'SAPI'];
  function claves(n) {
    return norm(n).split(/[^A-Z0-9]+/).filter(function (w) { return w.length >= 3 && VACIAS.indexOf(w) < 0; });
  }
  function parecidos(cat, nombre) {
    var c = compact(nombre);
    if (c.length < 3) return [];
    var k = claves(nombre);
    return state.cat[cat].filter(function (x) {
      var o = compact(x.nombre);
      if (o === c) return true;
      var kx = claves(x.nombre);
      if (k.length >= 2 && kx.length >= 2 && k[0] === kx[0] && k[1] === kx[1]) return true;
      if (c.length >= 5 && (o.indexOf(c) >= 0 || c.indexOf(o) >= 0)) return true;
      return c.length >= 6 && lev(o, c) <= 2;
    });
  }

  /* =========================================================================
     4. COMPONENTES
     ========================================================================= */

  // Las ventanas flotantes se colocan con position:fixed dentro del <dialog>
  // abierto (o del <body>), para que nunca queden recortadas.
  var popAbierto = null;
  function capaPara(el) { return el.closest('dialog[open]') || document.body; }
  function colocar(anchor, pop, opts) {
    opts = opts || {};
    var r = anchor.getBoundingClientRect();
    var vw = window.innerWidth, vh = window.innerHeight;
    var w = Math.max(r.width, opts.minWidth || 0);
    if (opts.width) w = opts.width;
    w = Math.min(w, vw - 16);
    var left = opts.alignRight ? r.right - w : r.left;
    left = Math.max(8, Math.min(left, vw - w - 8));
    pop.style.position = 'fixed';
    pop.style.left = left + 'px';
    pop.style.width = w + 'px';
    pop.style.right = 'auto';
    var abajo = vh - r.bottom, arriba = r.top;
    var alto = Math.min(pop.scrollHeight || 320, 360);
    if (abajo < alto + 12 && arriba > abajo) {
      pop.style.top = 'auto';
      pop.style.bottom = (vh - r.top + 6) + 'px';
    } else {
      pop.style.bottom = 'auto';
      pop.style.top = (r.bottom + 6) + 'px';
    }
  }
  function abrirPop(anchor, pop, opts, alCerrar) {
    cerrarPop();
    capaPara(anchor).appendChild(pop);
    colocar(anchor, pop, opts);
    function recolocar(e) { if (e && e.target && pop.contains(e.target)) return; colocar(anchor, pop, opts); }
    function fuera(e) { if (!pop.contains(e.target) && !anchor.contains(e.target)) cerrarPop(); }
    window.addEventListener('scroll', recolocar, true);
    window.addEventListener('resize', recolocar);
    document.addEventListener('mousedown', fuera, true);
    document.addEventListener('touchstart', fuera, true);
    popAbierto = {
      pop: pop,
      cerrar: function () {
        window.removeEventListener('scroll', recolocar, true);
        window.removeEventListener('resize', recolocar);
        document.removeEventListener('mousedown', fuera, true);
        document.removeEventListener('touchstart', fuera, true);
        pop.remove();
        popAbierto = null;
        if (alCerrar) alCerrar();
      },
      recolocar: function () { colocar(anchor, pop, opts); }
    };
    return popAbierto;
  }
  function cerrarPop() { if (popAbierto) popAbierto.cerrar(); }

  function resaltar(nombre, q) {
    if (!q) return esc(nombre);
    var n = norm(nombre), nq = norm(q), i = n.indexOf(nq);
    if (i < 0) return esc(nombre);
    return esc(nombre.slice(0, i)) + '<mark>' + esc(nombre.slice(i, i + q.length)) + '</mark>' + esc(nombre.slice(i + q.length));
  }

  /* ---------- 4.1 Lista seleccionable con búsqueda (cliente, proveedor, solicitante) ---------- */
  function Combo(host, opts) {
    var valor = opts.value == null ? null : opts.value;
    var cat = opts.catalog;
    host.innerHTML = '<button type="button" class="combo-btn" aria-haspopup="listbox" aria-expanded="false">' +
      '<span class="value"></span>' + icon('i-chev') + '</button>';
    var btn = host.firstChild, label = btn.firstChild;

    function pintar() {
      var it = valor != null ? state.byId[cat][valor] : null;
      if (it) { label.textContent = it.nombre + (it.activo ? '' : ' (inactivo)'); label.classList.remove('placeholder'); }
      else { label.textContent = opts.emptyLabel || 'Selecciona…'; label.classList.toggle('placeholder', !opts.emptyLabel); }
      btn.title = it ? it.nombre : '';
    }
    function opciones(q) {
      var nq = norm(q);
      var lista = state.cat[cat].filter(function (x) { return x.activo || x.id === valor || opts.includeInactive; });
      if (nq) {
        lista = lista.filter(function (x) { return norm(x.nombre).indexOf(nq) >= 0; })
          .sort(function (a, b) { return (norm(b.nombre).indexOf(nq) === 0) - (norm(a.nombre).indexOf(nq) === 0) || a.nombre.localeCompare(b.nombre); });
      }
      if (opts.emptyLabel && !nq) lista = [{ id: null, nombre: opts.emptyLabel, activo: true }].concat(lista);
      return lista;
    }
    function abrir(textoInicial) {
      if (popAbierto && popAbierto.pop._host === host) return;
      var pop = document.createElement('div');
      pop._host = host;
      pop.className = 'pop';
      pop.innerHTML = '<div class="pop-search"><input class="field" placeholder="Buscar…" aria-label="Buscar" autocomplete="off" spellcheck="false" /></div>' +
        '<ul class="pop-list" role="listbox"></ul>' + (opts.allowAdd ? '<div class="pop-foot" hidden><button type="button">' + icon('i-plus') + '<span></span></button></div>' : '');
      var input = $('input', pop), ul = $('ul', pop), foot = $('.pop-foot', pop);
      var activo = 0, lista = [];
      function render() {
        var q = input.value;
        lista = opciones(q);
        if (activo >= lista.length) activo = Math.max(0, lista.length - 1);
        ul.innerHTML = lista.length ? lista.map(function (x, i) {
          return '<li role="option" data-i="' + i + '" aria-selected="' + (x.id === valor) + '" class="' + (i === activo ? 'active' : '') + '">' +
            resaltar(x.nombre, x.id == null ? '' : q) + (x.activo ? '' : ' <span class="badge neutral plain">inactivo</span>') + '</li>';
        }).join('') : '<li class="none">Sin coincidencias</li>';
        if (foot) {
          var exacto = state.cat[cat].some(function (x) { return norm(x.nombre) === norm(q); });
          foot.hidden = !limpiar(q) || exacto;
          $('span', foot).textContent = 'Dar de alta “' + limpiar(q).toUpperCase() + '”';
        }
        var act = $('li.active', ul); if (act) act.scrollIntoView({ block: 'nearest' });
        if (popAbierto) popAbierto.recolocar();
      }
      function elegir(i) {
        var x = lista[i]; if (!x) return;
        valor = x.id; pintar(); cerrarPop(); btn.focus();
        if (opts.onChange) opts.onChange(valor);
      }
      ul.addEventListener('mousedown', function (e) { e.preventDefault(); });
      ul.addEventListener('click', function (e) { var li = e.target.closest('li[data-i]'); if (li) elegir(+li.dataset.i); });
      input.addEventListener('input', function () { activo = 0; render(); });
      input.addEventListener('keydown', function (e) {
        if (e.key === 'ArrowDown') { e.preventDefault(); activo = Math.min(activo + 1, lista.length - 1); render(); }
        else if (e.key === 'ArrowUp') { e.preventDefault(); activo = Math.max(activo - 1, 0); render(); }
        else if (e.key === 'Enter') { e.preventDefault(); if (lista.length) elegir(activo); else if (foot && !foot.hidden) $('button', foot).click(); }
        else if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); cerrarPop(); btn.focus(); }
        else if (e.key === 'Tab') { if (lista.length && limpiar(input.value)) elegir(activo); else cerrarPop(); }
      });
      if (foot) $('button', foot).addEventListener('click', async function () {
        var nombre = limpiar(input.value).toUpperCase();
        var par = parecidos(cat, nombre).filter(function (x) { return norm(x.nombre) !== norm(nombre); });
        cerrarPop();
        if (par.length) {
          var ok = await confirmar('¿Dar de alta “' + nombre + '”?',
            '<div class="notice warn">' + icon('i-alert') + '<span>Ya existen nombres parecidos: <b>' + par.slice(0, 5).map(function (x) { return esc(x.nombre); }).join('</b>, <b>') + '</b>. Si es el mismo, cancela y elígelo de la lista.</span></div>', 'Dar de alta');
          if (!ok) { btn.focus(); return; }
        }
        try {
          var nuevo = await altaCatalogo(cat, nombre);
          valor = nuevo.id; pintar(); btn.focus();
          if (opts.onChange) opts.onChange(valor);
        } catch (err) { toast('bad', 'No se pudo dar de alta', mensajeError(err)); }
      });
      abrirPop(btn, pop, { minWidth: 280 }, function () { btn.setAttribute('aria-expanded', 'false'); host.classList.remove('open'); });
      btn.setAttribute('aria-expanded', 'true');
      host.classList.add('open');
      input.value = textoInicial || '';
      var sel = opciones('').findIndex(function (x) { return x.id === valor; });
      activo = textoInicial ? 0 : Math.max(0, sel);
      render();
      input.focus();
    }
    btn.addEventListener('click', function () { if (host.classList.contains('open')) cerrarPop(); else abrir(); });
    btn.addEventListener('keydown', function (e) {
      if (e.key === 'ArrowDown' || e.key === 'Enter' || e.key === ' ') { e.preventDefault(); abrir(); }
      else if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) { e.preventDefault(); abrir(e.key); }
      else if ((e.key === 'Backspace' || e.key === 'Delete') && opts.emptyLabel) { valor = null; pintar(); if (opts.onChange) opts.onChange(null); }
    });
    document.addEventListener('catalogos', pintar);
    pintar();
    return {
      get: function () { return valor; },
      set: function (v) { valor = v == null ? null : v; pintar(); },
      focus: function () { btn.focus(); },
      el: btn
    };
  }

  /* ---------- 4.2 Fecha: se escribe (dd/mm/aaaa) o se elige en el calendario ---------- */
  function DateInput(host, opts) {
    opts = opts || {};
    host.innerHTML = '<input class="field" placeholder="dd/mm/aaaa" inputmode="numeric" maxlength="10" autocomplete="off" />' +
      '<button type="button" class="icon-btn" aria-label="Abrir calendario" title="Abrir calendario">' + icon('i-cal') + '</button>';
    var input = host.firstChild, btn = host.lastChild;
    if (opts.id) input.id = opts.id;

    function mascara(v) {
      v = v.replace(/[^\d\/]/g, '');
      if (v.indexOf('/') < 0) {
        var d = v.slice(0, 8);
        if (d.length > 4) return d.slice(0, 2) + '/' + d.slice(2, 4) + '/' + d.slice(4);
        if (d.length > 2) return d.slice(0, 2) + '/' + d.slice(2);
        return d;
      }
      var p = v.split('/').slice(0, 3);
      if (p.length > 1 && p[0].length === 1) p[0] = '0' + p[0];
      if (p.length > 2 && p[1].length === 1) p[1] = '0' + p[1];
      p[0] = p[0].slice(0, 2); if (p[1] != null) p[1] = p[1].slice(0, 2); if (p[2] != null) p[2] = p[2].slice(0, 4);
      return p.join('/');
    }
    input.addEventListener('input', function (e) {
      if (e.inputType && e.inputType.indexOf('insert') === 0) {
        var nuevo = mascara(input.value);
        if (nuevo !== input.value) input.value = nuevo;
      }
      if (opts.onInput) opts.onInput(get());
    });
    input.addEventListener('blur', function () {
      var v = get();
      if (v && v !== 'invalid') input.value = isoADMY(v);
      if (opts.onChange) opts.onChange(v);
    });
    input.addEventListener('keydown', function (e) {
      if ((e.altKey && e.key === 'ArrowDown') || e.key === 'F4') { e.preventDefault(); abrir(); }
      if (e.key === 'Enter') { var v = get(); if (v && v !== 'invalid') input.value = isoADMY(v); if (opts.onChange) opts.onChange(v); }
    });
    btn.addEventListener('click', function () { if (popAbierto && popAbierto.pop._host === host) cerrarPop(); else abrir(); });

    function get() { return dmyAISO(input.value); }
    function set(iso) { input.value = iso ? isoADMY(iso) : ''; }

    function abrir() {
      var actual = get();
      var base = actual && actual !== 'invalid' ? actual : hoyISO();
      var vy = +base.slice(0, 4), vm = +base.slice(5, 7), foco = base;
      var pop = document.createElement('div');
      pop._host = host;
      pop.className = 'pop cal';
      pop.setAttribute('role', 'dialog');
      pop.setAttribute('aria-label', 'Calendario');
      function render() {
        var min = opts.min ? opts.min() : null, max = opts.max ? opts.max() : null;
        var primero = new Date(Date.UTC(vy, vm - 1, 1)).getUTCDay(); // 0=domingo
        var inicio = (primero + 6) % 7; // semana inicia en lunes
        var diasMes = new Date(Date.UTC(vy, vm, 0)).getUTCDate();
        var hoy = hoyISO(), sel = get();
        var h = '<div class="cal-head"><button type="button" class="icon-btn sm plain" data-nav="-1" aria-label="Mes anterior">' + icon('i-left') + '</button>' +
          '<span class="cal-title">' + MESES[vm - 1] + ' ' + vy + '</span>' +
          '<button type="button" class="icon-btn sm plain" data-nav="1" aria-label="Mes siguiente">' + icon('i-right') + '</button></div><div class="cal-grid">';
        ['L', 'M', 'M', 'J', 'V', 'S', 'D'].forEach(function (d) { h += '<span class="dow">' + d + '</span>'; });
        for (var i = 0; i < 42; i++) {
          var dt = new Date(Date.UTC(vy, vm - 1, 1 - inicio + i));
          var iso = aISO(dt.getUTCFullYear(), dt.getUTCMonth() + 1, dt.getUTCDate());
          var cls = [];
          if (dt.getUTCMonth() !== vm - 1) cls.push('out');
          if (iso === hoy) cls.push('today');
          if (iso === sel) cls.push('sel');
          var dis = (min && iso < min) || (max && iso > max);
          h += '<button type="button" data-d="' + iso + '" class="' + cls.join(' ') + '"' + (dis ? ' disabled' : '') +
            ' tabindex="' + (iso === foco ? '0' : '-1') + '" aria-label="' + dt.getUTCDate() + ' de ' + MESES[dt.getUTCMonth()] + '">' + dt.getUTCDate() + '</button>';
          if (i >= 34 && i % 7 === 6 && 1 - inicio + i >= diasMes) break;
        }
        h += '</div><div class="cal-foot"><button type="button" class="btn-ghost btn-sm" data-hoy' + ((max && hoy > max) || (min && hoy < min) ? ' disabled' : '') + '>Hoy</button>' +
          (opts.clearable !== false ? '<button type="button" class="btn-ghost btn-sm" data-borrar>Borrar</button>' : '') + '</div>';
        pop.innerHTML = h;
        if (popAbierto) popAbierto.recolocar();
      }
      function elegir(iso) {
        input.value = iso ? isoADMY(iso) : '';
        cerrarPop(); input.focus();
        if (opts.onChange) opts.onChange(get());
      }
      pop.addEventListener('mousedown', function (e) { e.preventDefault(); });
      pop.addEventListener('click', function (e) {
        var b = e.target.closest('button'); if (!b || b.disabled) return;
        if (b.dataset.nav) { vm += +b.dataset.nav; if (vm < 1) { vm = 12; vy--; } if (vm > 12) { vm = 1; vy++; } render(); }
        else if (b.dataset.d) elegir(b.dataset.d);
        else if (b.hasAttribute('data-hoy')) elegir(hoyISO());
        else if (b.hasAttribute('data-borrar')) elegir(null);
      });
      pop.addEventListener('keydown', function (e) {
        var mov = { ArrowLeft: -1, ArrowRight: 1, ArrowUp: -7, ArrowDown: 7 }[e.key];
        if (mov) {
          e.preventDefault();
          var p = foco.split('-'), dt = new Date(Date.UTC(+p[0], p[1] - 1, +p[2] + mov));
          foco = aISO(dt.getUTCFullYear(), dt.getUTCMonth() + 1, dt.getUTCDate());
          vy = dt.getUTCFullYear(); vm = dt.getUTCMonth() + 1; render();
          var b = $('[data-d="' + foco + '"]', pop); if (b) b.focus();
        } else if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); cerrarPop(); input.focus(); }
      });
      abrirPop(host, pop, { width: 288, alignRight: true });
      render();
      var b = $('[data-d="' + foco + '"]', pop) || $('button.sel', pop);
      if (b) b.focus();
    }
    if (opts.value) set(opts.value);
    return { get: get, set: set, input: input, focus: function () { input.focus(); } };
  }

  /* ---------- 4.3 Autocompletado tipo Excel (tipo de servicio) ----------
     Mientras escribes, el resto de la palabra aparece seleccionado dentro del
     mismo campo. Enter o Tab lo aceptan; seguir escribiendo lo reemplaza. */
  function AutoExcel(input, opts) {
    var items = opts.items; // función → [{id, nombre}]
    var wrap = input.parentNode;
    var abierto = null;

    function coincidencias(txt) {
      var n = norm(txt);
      if (!n) return items();
      var empieza = [], contiene = [];
      items().forEach(function (x) {
        var nx = norm(x.nombre);
        if (nx.indexOf(n) === 0) empieza.push(x); else if (nx.indexOf(n) > 0) contiene.push(x);
      });
      return empieza.concat(contiene);
    }
    function exacto(txt) { var n = norm(txt); return items().filter(function (x) { return norm(x.nombre) === n; })[0] || null; }

    function lista(txt) {
      var ms = coincidencias(txt).slice(0, 8);
      if (!ms.length) { cerrar(); return; }
      if (!abierto) {
        var pop = document.createElement('div');
        pop.className = 'pop ac';
        pop._host = wrap;
        pop.innerHTML = '<ul class="pop-list" role="listbox"></ul>';
        pop.addEventListener('mousedown', function (e) { e.preventDefault(); });
        pop.addEventListener('click', function (e) {
          var li = e.target.closest('li[data-id]'); if (!li) return;
          aceptar(items().filter(function (x) { return String(x.id) === li.dataset.id; })[0]);
        });
        abierto = abrirPop(input, pop, {}, function () { abierto = null; });
      }
      $('ul', abierto.pop).innerHTML = ms.map(function (x, i) {
        return '<li role="option" data-id="' + x.id + '" class="' + (i === 0 ? 'active' : '') + '">' + resaltar(x.nombre, txt) + '</li>';
      }).join('');
      abierto.recolocar();
    }
    function cerrar() { if (abierto && popAbierto && popAbierto.pop === abierto.pop) cerrarPop(); abierto = null; }

    function aceptar(x) {
      if (!x) return;
      input.value = x.nombre;
      input.setSelectionRange(x.nombre.length, x.nombre.length);
      cerrar();
      if (opts.onCommit) opts.onCommit(x);
    }

    input.addEventListener('input', function (e) {
      var escrito = input.value;
      if (opts.onType) opts.onType();
      // Excel solo completa cuando se agrega texto al final (no al borrar)
      if (!e.inputType || e.inputType.indexOf('insert') !== 0 || input.selectionStart !== escrito.length) { lista(escrito); return; }
      var n = norm(escrito);
      if (!n) { lista(''); return; }
      var m = items().filter(function (x) { return norm(x.nombre).indexOf(n) === 0; })[0];
      if (m && m.nombre.length > escrito.length) {
        input.value = escrito + m.nombre.slice(escrito.length);
        input.setSelectionRange(escrito.length, input.value.length);
      }
      lista(escrito);
    });
    input.addEventListener('keydown', function (e) {
      var hayCompletado = input.selectionStart !== input.selectionEnd && input.selectionEnd === input.value.length;
      if (e.key === 'Enter') {
        e.preventDefault();
        var liAct = abierto ? $('li.active', abierto.pop) : null;
        var x = exacto(input.value) || (liAct && limpiar(input.value) ? items().filter(function (i) { return String(i.id) === liAct.dataset.id; })[0] : null);
        if (x) aceptar(x); else if (opts.onUnknown) { cerrar(); opts.onUnknown(limpiar(input.value)); }
      } else if (e.key === 'Tab') {
        var x2 = exacto(input.value);
        if (x2) aceptar(x2); else cerrar();
      } else if (e.key === 'Escape') {
        if (hayCompletado) { e.preventDefault(); e.stopPropagation(); input.value = input.value.slice(0, input.selectionStart); }
        else if (abierto) { e.preventDefault(); e.stopPropagation(); }
        cerrar();
      } else if ((e.key === 'ArrowDown' || e.key === 'ArrowUp') && abierto) {
        e.preventDefault();
        var lis = $$('li', abierto.pop), i = lis.findIndex(function (l) { return l.classList.contains('active'); });
        i = Math.max(0, Math.min(lis.length - 1, i + (e.key === 'ArrowDown' ? 1 : -1)));
        lis.forEach(function (l, k) { l.classList.toggle('active', k === i); });
        var it = items().filter(function (x) { return String(x.id) === lis[i].dataset.id; })[0];
        if (it) { input.value = it.nombre; input.setSelectionRange(it.nombre.length, it.nombre.length); }
      }
    });
    input.addEventListener('focus', function () { if (!input.value) lista(''); });
    input.addEventListener('blur', function () {
      setTimeout(function () {
        if (document.activeElement === input) return; // regresó al campo: no tocar lo que escribe
        if (document.activeElement && abierto && abierto.pop.contains(document.activeElement)) return;
        cerrar();
        var x = exacto(input.value);
        if (x) { if (input.value !== x.nombre) input.value = x.nombre; if (opts.onCommit) opts.onCommit(x); }
        else if (opts.onBlurUnknown) opts.onBlurUnknown(limpiar(input.value));
      }, 0);
    });
    return { valor: function () { return exacto(input.value); } };
  }

  /* =========================================================================
     4.4 COPIAR PARA CORREO
     Copia CONTRARRECIBO · NUR · CPT · DESCRIPCION · TIPO DE SERVICIO como
     tablita: al pegar en Outlook sale como tabla, y al pegar en Excel cada
     dato cae en su celda (el NUR se queda como texto).
     ========================================================================= */
  var COPIA_ENC = ['CONTRARRECIBO', 'NUR', 'CPT', 'DESCRIPCION', 'TIPO DE SERVICIO'];
  function copiaValores(r) { return [r.contrarrecibo || '', r.nur || '', r.cpt || '', r.descripcion || '', r.tipo_servicio || '']; }
  function tablaCorreo(rows) {
    var th = 'style="border:1px solid #808080;padding:3px 8px;background:#D9E1F2;font-weight:bold;text-align:left;white-space:nowrap"';
    var td = 'style="border:1px solid #808080;padding:3px 8px;mso-number-format:\'\\@\';white-space:nowrap"';
    var html = '<table style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:11pt">' +
      '<tr>' + COPIA_ENC.map(function (h) { return '<th ' + th + '>' + esc(h) + '</th>'; }).join('') + '</tr>' +
      rows.map(function (r) { return '<tr>' + copiaValores(r).map(function (v) { return '<td ' + td + '>' + esc(v) + '</td>'; }).join('') + '</tr>'; }).join('') +
      '</table>';
    var texto = COPIA_ENC.join('\t') + '\n' + rows.map(function (r) { return copiaValores(r).join('\t'); }).join('\n');
    return { html: html, texto: texto };
  }
  async function alPortapapeles(html, texto) {
    try {
      if (navigator.clipboard && window.ClipboardItem) {
        await navigator.clipboard.write([new ClipboardItem({
          'text/html': new Blob([html], { type: 'text/html' }),
          'text/plain': new Blob([texto], { type: 'text/plain' })
        })]);
        return true;
      }
    } catch (e) { /* se intenta el método alterno */ }
    var div = document.createElement('div');
    div.contentEditable = 'true';
    div.style.cssText = 'position:fixed;left:-9999px;top:0;opacity:0';
    div.innerHTML = html;
    capaPara(document.activeElement || document.body).appendChild(div);
    var rango = document.createRange(); rango.selectNodeContents(div);
    var sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(rango);
    var ok = false;
    try { ok = document.execCommand('copy'); } catch (e2) { ok = false; }
    sel.removeAllRanges(); div.remove();
    return ok;
  }
  async function copiarParaCorreo(rows) {
    if (!rows.length) return;
    var t = tablaCorreo(rows);
    var ok = await alPortapapeles(t.html, t.texto);
    if (ok) toast('ok', rows.length === 1 ? 'Copiado para correo' : rows.length + ' registros copiados para correo', 'Pega con Ctrl + V en tu correo o en Excel.');
    else toast('bad', 'No se pudo copiar', 'Tu navegador bloqueó el portapapeles. Inténtalo de nuevo.');
  }

  // Ícono y estilos de esta función (agregados desde aquí para no tocar otros archivos)
  (function () {
    var NS = 'http://www.w3.org/2000/svg', defs = document.querySelector('svg defs');
    if (defs && !document.getElementById('i-copy')) {
      var sym = document.createElementNS(NS, 'symbol');
      sym.setAttribute('id', 'i-copy'); sym.setAttribute('viewBox', '0 0 24 24'); sym.setAttribute('fill', 'none');
      sym.setAttribute('stroke', 'currentColor'); sym.setAttribute('stroke-width', '2'); sym.setAttribute('stroke-linecap', 'round'); sym.setAttribute('stroke-linejoin', 'round');
      var r1 = document.createElementNS(NS, 'rect'); r1.setAttribute('x', '8.5'); r1.setAttribute('y', '8.5'); r1.setAttribute('width', '12'); r1.setAttribute('height', '12'); r1.setAttribute('rx', '2.5');
      var p1 = document.createElementNS(NS, 'path'); p1.setAttribute('d', 'M15.5 8.5V5.5a2 2 0 0 0-2-2h-8a2 2 0 0 0-2 2v8a2 2 0 0 0 2 2h3');
      sym.appendChild(r1); sym.appendChild(p1); defs.appendChild(sym);
    }
    var st = document.createElement('style');
    st.textContent =
      '.grid .c-sel{width:36px;padding-left:14px;padding-right:4px}' +
      '.grid .c-sel input{width:16px;height:16px;margin:2px 0 0;accent-color:var(--accent);cursor:pointer}' +
      '.grid tr.is-selected td{background:var(--accent-soft)}' +
      '.grid tr.is-selected td.c-actions{background:color-mix(in srgb,var(--card-solid) 80%,var(--accent))}' +
      '.sel-bar{display:inline-flex;align-items:center;gap:6px;padding:3px 4px 3px 12px;border-radius:12px;background:var(--accent-soft);color:var(--accent-ink);font-size:12.5px;font-weight:650}' +
      '.sel-bar .btn-primary{min-height:30px;padding:5px 11px;font-size:12.5px;border-radius:9px}' +
      '.sesion{margin-top:18px;border:1px dashed var(--line-strong);border-radius:var(--r-md);padding:14px 16px}' +
      '.sesion-head{display:flex;align-items:center;gap:10px;flex-wrap:wrap}' +
      '.sesion-head b{font-size:13.5px}' +
      '.sesion table{width:100%;margin-top:10px;border-collapse:collapse;font-size:12.5px}' +
      '.sesion th{text-align:left;font-size:11px;color:var(--muted);padding:5px 8px;border-bottom:1px solid var(--line)}' +
      '.sesion td{padding:6px 8px;border-bottom:1px solid var(--line);font-variant-numeric:tabular-nums}' +
      '.sesion tr:last-child td{border-bottom:0}';
    document.head.appendChild(st);
  })();

  /* =========================================================================
     5. FORMULARIO DE AGREGADO (alta y edición)
     ========================================================================= */
  function Formulario(contenedor, cfg) {
    var modo = cfg.modo; // 'nuevo' | 'editar'
    var reg = cfg.registro || null; // fila de v_agregados al editar
    contenedor.innerHTML = '';
    var form = $('#tpl-form').content.firstElementChild.cloneNode(true);
    contenedor.appendChild(form);
    var uid = 'f' + Math.random().toString(36).slice(2, 8);
    function campo(n) { return $('[data-f="' + n + '"]', form); }
    function wrap(n) { return $('[data-wrap="' + n + '"]', form); }
    // Etiquetas accesibles
    $$('[data-wrap]', form).forEach(function (w) {
      var lbl = $('.label', w), f = $('input,select,textarea', w);
      if (lbl && f && f.tagName !== 'BUTTON') { f.id = uid + '-' + w.dataset.wrap; lbl.setAttribute('for', f.id); }
    });

    var original = reg ? {
      contrarrecibo: reg.contrarrecibo || '', nur: reg.nur || '', origen: reg.origen
    } : null;
    var tocado = {};
    var intentoGuardar = false;

    // --- Año surtido
    var selAnio = campo('anio_surtido');
    function pintarAnios() {
      var v = selAnio.value || (reg ? reg.anio_surtido : anioActual());
      var anios = state.anios.slice();
      if (reg && anios.indexOf(reg.anio_surtido) < 0) anios.push(reg.anio_surtido);
      selAnio.innerHTML = anios.sort(function (a, b) { return b - a; }).map(function (a) { return '<option value="' + a + '">' + a + '</option>'; }).join('');
      selAnio.value = String(v);
    }
    pintarAnios();

    // --- Listas seleccionables
    var combos = {};
    ['cliente_id', 'proveedor_id', 'solicitante_id'].forEach(function (k) {
      var host = campo(k);
      combos[k] = Combo(host, {
        catalog: host.dataset.catalog, allowAdd: true, value: reg ? reg[k] : null,
        onChange: function () { tocado[k] = true; validar(); }
      });
      $('.label', wrap(k)).addEventListener('click', function () { combos[k].focus(); });
    });

    // --- Tipo de servicio (autocompletado tipo Excel)
    var inTipo = campo('tipo_servicio');
    var tipoId = reg ? reg.tipo_servicio_id : null;
    if (reg) inTipo.value = nombreDe('tipos_servicio', tipoId);
    var acHint = $('.ac-hint', wrap('tipo_servicio_id'));
    var acHintTxt = acHint.innerHTML;
    function tipoDesconocido(txt) {
      if (!txt) { acHint.innerHTML = acHintTxt; return; }
      acHint.innerHTML = '“' + esc(txt.toUpperCase()) + '” no está en el catálogo. <button type="button" class="linkish">Dar de alta</button>';
      $('button', acHint).onclick = async function () {
        var par = parecidos('tipos_servicio', txt).filter(function (x) { return norm(x.nombre) !== norm(txt); });
        if (par.length) {
          var ok = await confirmar('¿Dar de alta “' + txt.toUpperCase() + '”?', '<div class="notice warn">' + icon('i-alert') + '<span>Ya existen tipos parecidos: <b>' + par.map(function (x) { return esc(x.nombre); }).join('</b>, <b>') + '</b>.</span></div>', 'Dar de alta');
          if (!ok) return;
        }
        try { var n = await altaCatalogo('tipos_servicio', txt); inTipo.value = n.nombre; tipoId = n.id; acHint.innerHTML = acHintTxt; validar(); }
        catch (e) { toast('bad', 'No se pudo dar de alta', mensajeError(e)); }
      };
    }
    AutoExcel(inTipo, {
      items: function () { return state.cat.tipos_servicio.filter(function (x) { return x.activo || x.id === tipoId; }); },
      onType: function () { tipoId = null; acHint.innerHTML = acHintTxt; },
      onCommit: function (x) { tipoId = x.id; tocado.tipo_servicio_id = true; acHint.innerHTML = acHintTxt; validar(); },
      onUnknown: function (txt) { tocado.tipo_servicio_id = true; tipoDesconocido(txt); validar(); },
      onBlurUnknown: function (txt) { tocado.tipo_servicio_id = true; tipoDesconocido(txt); validar(); }
    });

    // --- Textos
    var inContra = campo('contrarrecibo'), inNur = campo('nur'), inCpt = campo('cpt'), inDesc = campo('descripcion'), inNotas = campo('notas');
    if (reg) {
      inContra.value = reg.contrarrecibo || ''; inNur.value = reg.nur || ''; inCpt.value = reg.cpt || '';
      inDesc.value = reg.descripcion || ''; inNotas.value = reg.notas || '';
    }
    function soloDigitos(inp) {
      inp.addEventListener('input', function () {
        var p = inp.selectionStart, v = inp.value, d = v.replace(/\D/g, '');
        if (d !== v) { inp.value = d; inp.setSelectionRange(Math.max(0, p - (v.length - d.length)), Math.max(0, p - (v.length - d.length))); }
      });
    }
    // En registros migrados con NUR dañado, no forzamos dígitos hasta que lo cambien
    soloDigitos(inContra);
    if (!(reg && /[^\d]/.test(reg.nur || ''))) soloDigitos(inNur);
    function mayus(inp) {
      inp.addEventListener('input', function () {
        var p = inp.selectionStart, e2 = inp.selectionEnd, u = inp.value.toUpperCase();
        if (u !== inp.value) { inp.value = u; inp.setSelectionRange(p, e2); }
      });
    }
    mayus(inCpt); mayus(inDesc);
    [inContra, inNur, inCpt, inDesc, inNotas].forEach(function (inp) {
      inp.addEventListener('input', function () { if (intentoGuardar) validar(); actualizarVivo(); });
      // Un campo vacío solo se marca en rojo al intentar guardar (así no se mueve la pantalla al pasar por él)
      inp.addEventListener('blur', function () { if (limpiar(inp.value)) tocado[inp.dataset.f] = true; validar(); });
    });

    // --- Fechas
    var fSol = DateInput(campo('fecha_solicitud'), {
      value: reg ? reg.fecha_solicitud : (cfg.fechaInicial || hoyISO()), clearable: false,
      max: hoyISO, onChange: function () { tocado.fecha_solicitud = true; validar(); actualizarVivo(); }, onInput: function () { actualizarVivo(); }
    });
    var fCar = DateInput(campo('fecha_carga'), {
      value: reg ? reg.fecha_carga : null,
      min: function () { var s = fSol.get(); return s && s !== 'invalid' ? s : null; }, max: hoyISO,
      onChange: function () { tocado.fecha_carga = true; sincronizarEstatus(); validar(); actualizarVivo(); },
      onInput: function () { sincronizarEstatus(); actualizarVivo(); }
    });
    $('.label', wrap('fecha_solicitud')).setAttribute('for', fSol.input.id = uid + '-fsol');
    $('.label', wrap('fecha_carga')).setAttribute('for', fCar.input.id = uid + '-fcar');

    // --- Estatus: al poner fecha de carga queda CARGADO (como en la base)
    var selEst = campo('estatus');
    var estatusPrevio = reg && reg.estatus !== ESTATUS.CARG ? reg.estatus : ESTATUS.PEND;
    selEst.value = reg ? reg.estatus : ESTATUS.PEND;
    function sincronizarEstatus() {
      var fc = fCar.get();
      if (fc && fc !== 'invalid') {
        if (selEst.value !== ESTATUS.CARG) estatusPrevio = selEst.value;
        selEst.value = ESTATUS.CARG; selEst.disabled = true;
      } else {
        selEst.disabled = false;
        if (selEst.value === ESTATUS.CARG) {
          // Registros migrados que ya venían CARGADO sin fecha se respetan
          selEst.value = (reg && reg.estatus === ESTATUS.CARG && !reg.fecha_carga && !tocado.fecha_carga) ? ESTATUS.CARG : estatusPrevio;
        }
        selEst.querySelector('option[value="CARGADO"]').disabled = !(reg && reg.estatus === ESTATUS.CARG && !reg.fecha_carga);
      }
    }
    selEst.addEventListener('change', function () { estatusPrevio = selEst.value; actualizarVivo(); });
    sincronizarEstatus();

    // --- Estado en vivo
    var vivo = $('[data-live]', form);
    function actualizarVivo() {
      var fs = fSol.get(), fc = fCar.get();
      var e = estatusOps(fs && fs !== 'invalid' ? fs : null, selEst.value, fc && fc !== 'invalid' ? fc : null);
      vivo.innerHTML = '<span>Quedará como</span>' + badgeEstatus(selEst.value) + badgeOps(e.ops, null) +
        (e.dias != null ? '<span class="muted">' + e.dias + (e.dias === 1 ? ' día' : ' días') + ' desde la solicitud' +
          (e.ops === OPS.TIEMPO ? ' · pasa a “cambio de estatus” después de ' + DIAS_LIMITE + ' días' : '') + '</span>' : '');
    }
    actualizarVivo();

    // --- Revisión (registros migrados con observaciones)
    var revBox = $('[data-revision]', form), chkRev = null;
    if (reg && reg.requiere_revision) {
      revBox.hidden = false;
      revBox.innerHTML = '<div class="notice warn">' + icon('i-alert') + '<div><b>Este registro requiere revisión.</b> ' + esc(reg.motivo_revision || '') +
        '<label class="row mt-8" style="font-weight:600"><input type="checkbox" data-rev style="accent-color:var(--accent)"> Ya lo revisé, quitar la marca</label></div></div>';
      chkRev = $('[data-rev]', revBox);
    }

    // --- NUR repetido (aviso, no bloquea: hay repetidos legítimos)
    var nurAviso = null;
    var revisarNur = debounce(async function () {
      var v = limpiar(inNur.value);
      quitarNota('nur');
      nurAviso = null;
      if (!v || (reg && v === reg.nur)) return;
      var r = await sb.from('v_agregados').select('id,cliente,proveedor,cpt,descripcion,fecha_solicitud,estatus').eq('nur', v).eq('eliminado', false).limit(5);
      if (r.error || !r.data.length) return;
      var otros = r.data.filter(function (x) { return !reg || x.id !== reg.id; });
      if (!otros.length) return;
      nurAviso = otros;
      nota('nur', 'warn', 'Este NUR ya está registrado ' + otros.length + (otros.length === 1 ? ' vez' : ' veces') + ': ' +
        otros.slice(0, 2).map(function (x) { return esc(x.cpt + ' · ' + x.descripcion + ' (' + isoADMY(x.fecha_solicitud) + ')'); }).join('; ') + '.');
    }, 350);
    inNur.addEventListener('input', revisarNur);

    // --- CPT: sugiere la descripción usada antes con ese CPT
    var revisarCpt = debounce(async function () {
      var v = limpiar(inCpt.value);
      quitarNota('cpt');
      if (!v) return;
      var r = await sb.from('agregados').select('descripcion').eq('cpt', v).order('created_at', { ascending: false }).limit(1);
      if (r.error || !r.data.length) return;
      var d = r.data[0].descripcion;
      if (!limpiar(inDesc.value)) {
        inDesc.value = d; validar();
        nota('cpt', 'info', 'Descripción tomada de un registro anterior con este CPT.');
      } else if (norm(inDesc.value) !== norm(d)) {
        nota('cpt', 'info', 'Antes se usó: “' + esc(d) + '”. <button type="button" class="linkish" data-usar>Usar esta</button>');
        var b = $('[data-usar]', wrap('cpt'));
        if (b) b.onclick = function () { inDesc.value = d; quitarNota('cpt'); validar(); };
      }
    }, 400);
    inCpt.addEventListener('input', revisarCpt);

    function nota(campoN, tipo, html) {
      var w = wrap(campoN); quitarNota(campoN);
      var p = document.createElement('p');
      p.className = 'field-note'; p.dataset.nota = '1';
      p.style.color = tipo === 'warn' ? 'var(--warn-ink)' : '';
      p.innerHTML = html;
      w.appendChild(p);
    }
    function quitarNota(campoN) { var n = $('[data-nota]', wrap(campoN)); if (n) n.remove(); }

    // --- Validación: ningún dato obligatorio en blanco
    function errores() {
      var e = {};
      var hoy = hoyISO();
      if (!selAnio.value) e.anio_surtido = 'Elige el año.';
      if (!combos.cliente_id.get()) e.cliente_id = 'Elige el cliente.';
      if (!combos.proveedor_id.get()) e.proveedor_id = 'Elige el proveedor.';
      if (!combos.solicitante_id.get()) e.solicitante_id = 'Elige quién lo solicita.';
      if (!tipoId) e.tipo_servicio_id = limpiar(inTipo.value) ? 'Elige un tipo de la lista (o dalo de alta).' : 'Escribe el tipo de servicio.';
      var contra = limpiar(inContra.value);
      var contraOpcional = reg && reg.origen === 'EXCEL_2026' && !original.contrarrecibo;
      if (!contra && !contraOpcional) e.contrarrecibo = 'Escribe el contrarrecibo.';
      else if (contra && !/^\d+$/.test(contra)) e.contrarrecibo = 'Solo números.';
      var nur = limpiar(inNur.value);
      if (!nur) e.nur = 'Escribe el NUR.';
      else if (!/^\d+$/.test(nur) && !(reg && nur === original.nur)) e.nur = 'Solo números.';
      if (!limpiar(inCpt.value)) e.cpt = 'Escribe el CPT.';
      if (!limpiar(inDesc.value)) e.descripcion = 'Escribe la descripción.';
      var fs = fSol.get();
      if (!fs) e.fecha_solicitud = 'Escribe o elige la fecha.';
      else if (fs === 'invalid') e.fecha_solicitud = 'Fecha no válida. Usa dd/mm/aaaa.';
      else if (fs > hoy) e.fecha_solicitud = 'No puede ser una fecha futura.';
      var fc = fCar.get();
      if (fc === 'invalid') e.fecha_carga = 'Fecha no válida. Usa dd/mm/aaaa.';
      else if (fc && fc > hoy) e.fecha_carga = 'No puede ser una fecha futura.';
      else if (fc && fs && fs !== 'invalid' && fc < fs) e.fecha_carga = 'No puede ser anterior a la fecha de solicitud.';
      return e;
    }
    function validar(forzar) {
      var e = errores();
      $$('[data-wrap]', form).forEach(function (w) {
        var k = w.dataset.wrap, msg = e[k];
        var mostrarErr = msg && (forzar || intentoGuardar || tocado[k] || tocado[k.replace('_id', '')]);
        w.classList.toggle('is-invalid', !!mostrarErr);
        var p = $('.field-error', w);
        if (mostrarErr) { if (!p) { p = document.createElement('p'); p.className = 'field-error'; var ref = $('.ac-hint', w); w.insertBefore(p, ref || null); } p.textContent = msg; }
        else if (p) p.remove();
      });
      if (contraNota) contraNota();
      if (intentoGuardar && !forzar) {
        var n = Object.keys(e).length;
        if (!n) errBox.hidden = true;
        else if (!errBox.hidden && /Faltan/.test(errBox.textContent)) $('b', errBox).textContent = 'Faltan ' + n + (n === 1 ? ' dato.' : ' datos.');
      }
      return e;
    }
    var contraNota = null;
    if (reg && reg.origen === 'EXCEL_2026' && !reg.contrarrecibo) {
      contraNota = function () { if (!$('[data-nota]', wrap('contrarrecibo')) && !limpiar(inContra.value)) nota('contrarrecibo', 'info', 'Venía vacío en el Excel; puedes completarlo.'); };
    }

    // --- Botones
    var acciones = $('[data-actions]', form), errBox = $('[data-form-error]', form);
    acciones.innerHTML = (modo === 'nuevo'
      ? '<button type="button" class="btn-ghost" data-reset>Limpiar</button><button type="submit" class="btn-primary">' + icon('i-check') + 'Guardar agregado</button>'
      : '<span class="muted grow" style="font-size:12.5px">' + (reg.origen === 'EXCEL_2026' ? 'Migrado del Excel (fila ' + reg.fila_excel + ')' : 'Capturado por ' + esc(reg.created_by || '—')) + ' · última modificación ' + esc(fechaHora(reg.updated_at)) + (reg.updated_by ? ' por ' + esc(reg.updated_by) : '') + '</span>' +
        '<button type="button" class="btn-ghost" data-cancel>Cancelar</button><button type="submit" class="btn-primary">' + icon('i-check') + 'Guardar cambios</button>');

    form.addEventListener('submit', async function (ev) {
      ev.preventDefault();
      if (popAbierto) cerrarPop();
      intentoGuardar = true;
      // El tipo de servicio se resuelve al salir del campo; si quedó escrito exacto, se toma
      if (!tipoId) { var ex = state.cat.tipos_servicio.filter(function (x) { return norm(x.nombre) === norm(inTipo.value); })[0]; if (ex) { tipoId = ex.id; inTipo.value = ex.nombre; } }
      var e = validar(true);
      errBox.hidden = true;
      var claves = Object.keys(e);
      if (claves.length) {
        errBox.hidden = false;
        errBox.innerHTML = '<div class="notice bad">' + icon('i-alert') + '<span><b>Faltan ' + claves.length + (claves.length === 1 ? ' dato' : ' datos') + '.</b> Revisa los campos marcados en rojo.</span></div>';
        var primero = $('.is-invalid', form);
        if (primero) { var f = $('input,button,select,textarea', primero); if (f) f.focus(); }
        return;
      }
      var fc = fCar.get();
      var datos = {
        anio_surtido: +selAnio.value,
        cliente_id: combos.cliente_id.get(),
        proveedor_id: combos.proveedor_id.get(),
        solicitante_id: combos.solicitante_id.get(),
        tipo_servicio_id: tipoId,
        contrarrecibo: limpiar(inContra.value) || null,
        nur: limpiar(inNur.value),
        cpt: limpiar(inCpt.value).toUpperCase(),
        descripcion: limpiar(inDesc.value).toUpperCase(),
        fecha_solicitud: fSol.get(),
        fecha_carga: fc || null,
        estatus: fc ? ESTATUS.CARG : selEst.value,
        notas: limpiar(inNotas.value) || null
      };
      if (chkRev && chkRev.checked) { datos.requiere_revision = false; }

      // NUR repetido: confirmar
      var rep = await sb.from('v_agregados').select('id,cpt,descripcion,fecha_solicitud,proveedor').eq('nur', datos.nur).eq('eliminado', false).limit(5);
      var otros = (rep.data || []).filter(function (x) { return !reg || x.id !== reg.id; });
      if (otros.length && !(reg && datos.nur === reg.nur)) {
        var ok = await confirmar('Este NUR ya existe',
          '<p class="card-hint" style="margin-bottom:10px">El NUR <b>' + esc(datos.nur) + '</b> ya está registrado:</p>' +
          '<ul style="margin:0;padding-left:18px;font-size:13px">' + otros.map(function (x) { return '<li>' + esc(x.cpt + ' · ' + x.descripcion) + ' — ' + esc(x.proveedor) + ', ' + isoADMY(x.fecha_solicitud) + '</li>'; }).join('') + '</ul>' +
          '<p class="card-hint" style="margin-top:10px">¿Es otro estudio del mismo paciente? Si es así, guárdalo.</p>', 'Guardar de todas formas');
        if (!ok) { inNur.focus(); return; }
      }

      var btn = $('button[type="submit"]', form);
      cargando(btn, true);
      var r = modo === 'nuevo'
        ? await sb.from('agregados').insert(datos).select('id').single()
        : await sb.from('agregados').update(datos).eq('id', reg.id).select('id').single();
      cargando(btn, false);
      if (r.error) {
        errBox.hidden = false;
        errBox.innerHTML = '<div class="notice bad">' + icon('i-alert') + '<span><b>No se guardó.</b> ' + esc(mensajeError(r.error)) + '</span></div>';
        return;
      }
      state.consultaSucia = true;
      actualizarContadorNav();
      if (state.anios.indexOf(datos.anio_surtido) < 0) cargarCatalogos();
      if (modo === 'nuevo') {
        toast('ok', 'Agregado guardado', 'NUR ' + datos.nur + ' · ' + datos.cpt + (datos.fecha_carga ? ' · cargado' : ''));
        if (cfg.onCreado) cfg.onCreado({ contrarrecibo: datos.contrarrecibo, nur: datos.nur, cpt: datos.cpt, descripcion: datos.descripcion, tipo_servicio: nombreDe('tipos_servicio', datos.tipo_servicio_id) });
        limpiarParaSiguiente();
      } else {
        toast('ok', 'Cambios guardados', 'NUR ' + datos.nur);
        if (cfg.onGuardado) cfg.onGuardado(r.data.id);
      }
    });

    function limpiarParaSiguiente() {
      // Se conservan cliente, proveedor, solicitante, tipo, contrarrecibo, año y fecha de solicitud
      inNur.value = ''; inCpt.value = ''; inDesc.value = ''; inNotas.value = '';
      fCar.set(null); sincronizarEstatus(); selEst.value = ESTATUS.PEND; estatusPrevio = ESTATUS.PEND;
      tocado = {}; intentoGuardar = false;
      ['nur', 'cpt'].forEach(quitarNota);
      validar(); actualizarVivo(); errBox.hidden = true;
      inNur.focus();
    }
    var bReset = $('[data-reset]', acciones);
    if (bReset) bReset.addEventListener('click', function () {
      Formulario(contenedor, cfg);
      $('.combo-btn', contenedor).focus();
    });
    var bCancel = $('[data-cancel]', acciones);
    if (bCancel) bCancel.addEventListener('click', function () { if (cfg.onCancelar) cfg.onCancelar(); });

    // Enter en un campo de texto pasa al siguiente (como en Excel), no envía
    form.addEventListener('keydown', function (e) {
      if (e.key !== 'Enter' || e.defaultPrevented || e.target.tagName === 'TEXTAREA' || e.target.type === 'submit' || e.target.tagName === 'BUTTON') return;
      if (e.ctrlKey || e.metaKey) { e.preventDefault(); form.requestSubmit(); return; }
      e.preventDefault();
      var focos = $$('input:not([disabled]):not([type=checkbox]), select:not([disabled]), textarea, .combo-btn', form).filter(function (x) { return x.offsetParent !== null; });
      var i = focos.indexOf(e.target);
      if (i >= 0 && focos[i + 1]) focos[i + 1].focus();
    });

    document.addEventListener('catalogos', pintarAnios);
    return { form: form, focus: function () { combos.cliente_id.focus(); } };
  }

  /* =========================================================================
     6. CONSULTA
     ========================================================================= */
  var COLS = [
    { k: 'anio_surtido', t: 'Año', cls: 'c-num' },
    { k: 'cliente', t: 'Cliente' },
    { k: 'proveedor', t: 'Proveedor', cls: 'c-prov' },
    { k: 'contrarrecibo', t: 'Contrarrecibo', cls: 'c-num' },
    { k: 'nur', t: 'NUR', cls: 'c-nur' },
    { k: 'cpt', t: 'CPT', cls: 'c-num' },
    { k: 'descripcion', t: 'Descripción', cls: 'c-desc' },
    { k: 'tipo_servicio', t: 'Tipo de servicio' },
    { k: 'solicitante', t: 'Solicitante' },
    { k: 'fecha_solicitud', t: 'Fecha solicitud', cls: 'c-date' },
    { k: 'fecha_carga', t: 'Fecha de carga', cls: 'c-date' },
    { k: 'estatus', t: 'Estatus' },
    { k: 'estatus_operaciones', t: 'Estatus operaciones' },
    { k: 'notas', t: 'Notas', cls: 'c-notes' }
  ];
  var Q = {
    anio: null, q: '', cliente_id: null, proveedor_id: null, tipo_servicio_id: null, solicitante_id: null,
    estatus: '', ops: '', desde: null, hasta: null, revision: false, bajas: false,
    orden: 'fecha_solicitud', asc: false, pagina: 0, tam: 50
  };
  var filas = [];
  var filtros = {};
  var pedido = 0;

  function iniciarConsulta() {
    Q.anio = anioActual();
    // Encabezados ordenables
    $('#grid-head').innerHTML = '<th class="c-sel"><input type="checkbox" id="sel-all" title="Seleccionar los de esta página" aria-label="Seleccionar todos los de esta página"></th>' + COLS.map(function (c) {
      return '<th class="' + (c.cls || '') + '" data-k="' + c.k + '"><button type="button">' + c.t + ' <span class="sort">↕</span></button></th>';
    }).join('') + '<th class="c-actions"><span class="sr-only">Acciones</span></th>';
    $('#grid-head').addEventListener('click', function (e) {
      var th = e.target.closest('th[data-k]'); if (!th) return;
      if (Q.orden === th.dataset.k) Q.asc = !Q.asc; else { Q.orden = th.dataset.k; Q.asc = true; }
      Q.pagina = 0; cargarTabla();
    });

    pintarFiltroAnios();
    document.addEventListener('catalogos', pintarFiltroAnios);
    $('#f-anio').addEventListener('change', function () { Q.anio = this.value ? +this.value : null; Q.pagina = 0; recargar(); });
    $('#f-q').addEventListener('input', debounce(function () { Q.q = this.value; Q.pagina = 0; cargarTabla(); }, 300));
    $('#f-estatus').addEventListener('change', function () { Q.estatus = this.value; Q.pagina = 0; cargarTabla(); });
    $('#f-ops').addEventListener('change', function () { Q.ops = this.value; Q.pagina = 0; pintarKpiActivo(); cargarTabla(); });
    $('#f-revision').addEventListener('change', function () { Q.revision = this.checked; Q.pagina = 0; cargarTabla(); });
    $('#f-bajas').addEventListener('change', function () { Q.bajas = this.checked; Q.pagina = 0; cargarTabla(); });
    $$('.combo[data-filter]').forEach(function (host) {
      var k = host.dataset.filter;
      filtros[k] = Combo(host, { catalog: host.dataset.catalog, emptyLabel: 'Todos', includeInactive: true, onChange: function (v) { Q[k] = v; Q.pagina = 0; cargarTabla(); } });
    });
    filtros.desde = DateInput($('[data-date="f-desde"]'), { onChange: function (v) { Q.desde = v && v !== 'invalid' ? v : null; Q.pagina = 0; cargarTabla(); } });
    filtros.hasta = DateInput($('[data-date="f-hasta"]'), { onChange: function (v) { Q.hasta = v && v !== 'invalid' ? v : null; Q.pagina = 0; cargarTabla(); } });
    $('#btn-more').addEventListener('click', function () {
      var fm = $('#filters-more'); fm.hidden = !fm.hidden;
      $('span', this).textContent = fm.hidden ? 'Más filtros' : 'Menos filtros';
    });
    $('#btn-clear').addEventListener('click', limpiarFiltros);
    $('#btn-refresh').addEventListener('click', recargar);
    $('#btn-export').addEventListener('click', exportarConsulta);
    $('#page-size').addEventListener('change', function () { Q.tam = +this.value; Q.pagina = 0; cargarTabla(); });
    $('#pg-prev').addEventListener('click', function () { if (Q.pagina > 0) { Q.pagina--; cargarTabla(); } });
    $('#pg-next').addEventListener('click', function () { Q.pagina++; cargarTabla(); });
    $$('.kpi').forEach(function (b) {
      b.addEventListener('click', function () {
        var v = b.dataset.kpi === 'todos' ? '' : b.dataset.kpi;
        Q.ops = Q.ops === v ? '' : v;
        $('#f-ops').value = Q.ops;
        if (Q.ops) $('#filters-more').hidden = false, $('#btn-more span').textContent = 'Menos filtros';
        Q.pagina = 0; pintarKpiActivo(); cargarTabla();
      });
    });
    $('#grid-body').addEventListener('click', accionFila);
  }

  function pintarFiltroAnios() {
    var sel = $('#f-anio'); if (!sel) return;
    sel.innerHTML = '<option value="">Todos los años</option>' + state.anios.map(function (a) { return '<option value="' + a + '">' + a + '</option>'; }).join('');
    sel.value = Q.anio == null ? '' : String(Q.anio);
  }
  function pintarKpiActivo() {
    $$('.kpi').forEach(function (b) { b.setAttribute('aria-pressed', String((b.dataset.kpi === 'todos' ? '' : b.dataset.kpi) === Q.ops && !!Q.ops)); });
  }
  function hayFiltros() {
    return !!(Q.q || Q.cliente_id || Q.proveedor_id || Q.tipo_servicio_id || Q.solicitante_id || Q.estatus || Q.ops || Q.desde || Q.hasta || Q.revision || Q.bajas);
  }
  function limpiarFiltros() {
    Q.q = ''; Q.cliente_id = Q.proveedor_id = Q.tipo_servicio_id = Q.solicitante_id = null;
    Q.estatus = Q.ops = ''; Q.desde = Q.hasta = null; Q.revision = Q.bajas = false; Q.pagina = 0;
    $('#f-q').value = ''; $('#f-estatus').value = ''; $('#f-ops').value = '';
    $('#f-revision').checked = false; $('#f-bajas').checked = false;
    ['cliente_id', 'proveedor_id', 'tipo_servicio_id', 'solicitante_id'].forEach(function (k) { filtros[k].set(null); });
    filtros.desde.set(null); filtros.hasta.set(null);
    pintarKpiActivo(); cargarTabla();
  }

  function aplicarFiltros(q, opts) {
    opts = opts || {};
    if (!Q.bajas) q = q.eq('eliminado', false);
    if (Q.anio) q = q.eq('anio_surtido', Q.anio);
    ['cliente_id', 'proveedor_id', 'tipo_servicio_id', 'solicitante_id'].forEach(function (k) { if (Q[k]) q = q.eq(k, Q[k]); });
    if (Q.estatus) q = q.eq('estatus', Q.estatus);
    if (Q.ops && !opts.sinOps) q = q.eq('estatus_operaciones', Q.ops);
    if (Q.desde) q = q.gte('fecha_solicitud', Q.desde);
    if (Q.hasta) q = q.lte('fecha_solicitud', Q.hasta);
    if (Q.revision) q = q.eq('requiere_revision', true);
    var s = limpiar(Q.q).replace(/[,()*%:"\\]/g, ' ').trim();
    if (s) {
      var pat = '%' + s.replace(/\s+/g, '%') + '%';
      q = q.or(['nur', 'cpt', 'descripcion', 'contrarrecibo', 'proveedor', 'notas'].map(function (c) { return c + '.ilike.' + pat; }).join(','));
    }
    return q;
  }

  async function cargarTabla() {
    var mio = ++pedido;
    var body = $('#grid-body');
    $('#btn-clear').hidden = !hayFiltros();
    $$('#grid-head th[data-k]').forEach(function (th) {
      var on = th.dataset.k === Q.orden;
      if (on) th.setAttribute('aria-sort', Q.asc ? 'ascending' : 'descending'); else th.removeAttribute('aria-sort');
      $('.sort', th).textContent = on ? (Q.asc ? '↑' : '↓') : '↕';
    });
    if (!filas.length) body.innerHTML = Array(8).join('<tr class="skeleton">' + Array(COLS.length + 3).join('<td><div></div></td>') + '</tr>');
    var desde = Q.pagina * Q.tam;
    var q = aplicarFiltros(sb.from('v_agregados').select('*', { count: 'exact' }))
      .order(Q.orden, { ascending: Q.asc, nullsFirst: false })
      .order('id', { ascending: false })
      .range(desde, desde + Q.tam - 1);
    var r = await q;
    if (mio !== pedido) return;
    if (r.error) {
      body.innerHTML = '<tr><td colspan="' + (COLS.length + 2) + '"><div class="empty"><strong>No se pudo consultar</strong>' + esc(mensajeError(r.error)) + '</div></td></tr>';
      return;
    }
    filas = r.data;
    var total = r.count || 0;
    if (!filas.length && Q.pagina > 0 && total > 0) { Q.pagina = Math.max(0, Math.ceil(total / Q.tam) - 1); return cargarTabla(); }
    pintarFilas();
    pintarSeleccion();
    var paginas = Math.max(1, Math.ceil(total / Q.tam));
    $('#pager-info').textContent = total ? ('Mostrando ' + (desde + 1).toLocaleString('es-MX') + '–' + (desde + filas.length).toLocaleString('es-MX') + ' de ' + total.toLocaleString('es-MX')) : 'Sin registros';
    $('#pg-label').textContent = (Q.pagina + 1) + ' / ' + paginas;
    $('#pg-prev').disabled = Q.pagina === 0;
    $('#pg-next').disabled = Q.pagina >= paginas - 1;
    state.consultaSucia = false;
  }

  // Selección para copiar varios a la vez (se conserva al cambiar de página o de búsqueda)
  var seleccion = new Map();
  function pintarSeleccion() {
    var bar = $('#sel-bar');
    if (!bar) {
      bar = document.createElement('span');
      bar.id = 'sel-bar'; bar.className = 'sel-bar'; bar.hidden = true;
      bar.innerHTML = '<span data-n></span><button type="button" class="btn-primary" data-copiar>' + icon('i-copy') + 'Copiar para correo</button>' +
        '<button type="button" class="icon-btn sm plain" data-limpiar-sel title="Quitar selección" aria-label="Quitar selección">' + icon('i-x') + '</button>';
      $('#btn-refresh').parentNode.insertBefore(bar, $('#btn-refresh'));
      $('[data-copiar]', bar).addEventListener('click', function () { copiarParaCorreo(Array.from(seleccion.values())); });
      $('[data-limpiar-sel]', bar).addEventListener('click', function () { seleccion.clear(); pintarFilas(); pintarSeleccion(); });
    }
    var n = seleccion.size;
    bar.hidden = !n;
    $('[data-n]', bar).textContent = n + (n === 1 ? ' seleccionado' : ' seleccionados');
    var todos = $('#sel-all');
    if (todos) {
      var enPagina = filas.filter(function (r) { return seleccion.has(r.id); }).length;
      todos.checked = filas.length > 0 && enPagina === filas.length;
      todos.indeterminate = enPagina > 0 && enPagina < filas.length;
    }
  }
  document.addEventListener('change', function (e) {
    var t = e.target;
    if (t.id === 'sel-all') {
      filas.forEach(function (r) { if (t.checked) seleccion.set(r.id, r); else seleccion.delete(r.id); });
      pintarFilas(); pintarSeleccion();
    } else if (t.dataset && t.dataset.sel != null && t.closest('#grid-body')) {
      var r = filas[+t.dataset.sel]; if (!r) return;
      if (t.checked) seleccion.set(r.id, r); else seleccion.delete(r.id);
      t.closest('tr').classList.toggle('is-selected', t.checked);
      pintarSeleccion();
    }
  });

  function pintarFilas() {
    var body = $('#grid-body');
    if (!filas.length) {
      body.innerHTML = '<tr><td colspan="' + (COLS.length + 2) + '"><div class="empty"><strong>No hay registros con estos filtros</strong>' +
        (hayFiltros() ? '<button type="button" class="linkish" data-limpiar>Limpiar filtros</button>' : '<a class="linkish" href="#nuevo">Capturar un agregado</a>') + '</div></td></tr>';
      var bl = $('[data-limpiar]', body); if (bl) bl.onclick = limpiarFiltros;
      return;
    }
    body.innerHTML = filas.map(function (r, i) {
      var flag = r.requiere_revision ? '<span class="flag" title="' + esc('Requiere revisión: ' + (r.motivo_revision || '')) + '">' + icon('i-flag') + '</span>' : '';
      var acc = r.eliminado
        ? '<button class="icon-btn sm plain" data-a="hist" title="Historial" aria-label="Historial">' + icon('i-history') + '</button>' +
          '<button class="icon-btn sm plain ok" data-a="reactivar" title="Reactivar" aria-label="Reactivar">' + icon('i-undo') + '</button>'
        : (r.estatus !== ESTATUS.CARG ? '<button class="icon-btn sm plain ok" data-a="carga" title="Capturar fecha de carga" aria-label="Capturar fecha de carga">' + icon('i-truck') + '</button>' : '') +
          '<button class="icon-btn sm plain" data-a="copiar" title="Copiar para correo" aria-label="Copiar para correo">' + icon('i-copy') + '</button>' +
          '<button class="icon-btn sm plain" data-a="editar" title="Editar" aria-label="Editar">' + icon('i-edit') + '</button>' +
          '<button class="icon-btn sm plain" data-a="hist" title="Historial" aria-label="Historial">' + icon('i-history') + '</button>' +
          '<button class="icon-btn sm plain bad" data-a="baja" title="Dar de baja" aria-label="Dar de baja">' + icon('i-trash') + '</button>';
      var marcado = seleccion.has(r.id);
      var clases = (r.eliminado ? 'is-deleted' : '') + (marcado ? ' is-selected' : '');
      return '<tr data-i="' + i + '" class="' + clases.trim() + '"' + (r.eliminado ? ' title="' + esc('Baja: ' + (r.motivo_baja || '')) + '"' : '') + '>' +
        '<td class="c-sel"><input type="checkbox" data-sel="' + i + '"' + (marcado ? ' checked' : '') + ' aria-label="Seleccionar NUR ' + esc(r.nur) + '"></td>' +
        '<td class="c-num">' + r.anio_surtido + '</td>' +
        '<td>' + esc(r.cliente) + '</td>' +
        '<td class="c-prov">' + esc(r.proveedor) + '</td>' +
        '<td class="c-num">' + esc(r.contrarrecibo || '') + '</td>' +
        '<td class="c-nur">' + esc(r.nur) + flag + '</td>' +
        '<td class="c-num">' + esc(r.cpt) + '</td>' +
        '<td class="c-desc">' + esc(r.descripcion) + '</td>' +
        '<td>' + esc(r.tipo_servicio) + '</td>' +
        '<td>' + esc(r.solicitante) + '</td>' +
        '<td class="c-date">' + isoADMY(r.fecha_solicitud) + '</td>' +
        '<td class="c-date">' + (r.fecha_carga ? isoADMY(r.fecha_carga) : '<span class="muted">—</span>') + '</td>' +
        '<td>' + badgeEstatus(r.estatus) + '</td>' +
        '<td>' + badgeOps(r.estatus_operaciones, r.dias_transcurridos) + '</td>' +
        '<td class="c-notes">' + esc(r.notas || '') + '</td>' +
        '<td class="c-actions"><div class="row-actions">' + acc + '</div></td></tr>';
    }).join('');
  }

  async function cargarKpis() {
    function cuenta(ops) {
      var q = sb.from('v_agregados').select('id', { count: 'exact', head: true }).eq('eliminado', false);
      if (Q.anio) q = q.eq('anio_surtido', Q.anio);
      if (ops) q = q.eq('estatus_operaciones', ops);
      return q;
    }
    var r = await Promise.all([cuenta(null), cuenta(OPS.OK), cuenta(OPS.TIEMPO), cuenta(OPS.CAMBIO)]);
    if (r.some(function (x) { return x.error; })) return;
    var n = r.map(function (x) { return x.count || 0; });
    $('#k-total').textContent = n[0].toLocaleString('es-MX');
    $('#k-ok').textContent = n[1].toLocaleString('es-MX');
    $('#k-warn').textContent = n[2].toLocaleString('es-MX');
    $('#k-bad').textContent = n[3].toLocaleString('es-MX');
    var t = n[0] || 1;
    $('#m-ok').style.width = (n[1] / t * 100) + '%';
    $('#m-warn').style.width = (n[2] / t * 100) + '%';
    $('#m-bad').style.width = (n[3] / t * 100) + '%';
    $$('.kpi-label', $('.kpis'))[0].textContent = Q.anio ? 'Registros ' + Q.anio : 'Registros (todos los años)';
  }

  async function actualizarContadorNav() {
    var r = await sb.from('v_pendientes_vencidos').select('id', { count: 'exact', head: true });
    if (r.error) return;
    var n = r.count || 0;
    var el = $('#nav-vencidos');
    el.hidden = !n; el.textContent = n; el.title = n + ' solicitudes con más de ' + DIAS_LIMITE + ' días sin cargar';
    var ac = $('#alert-count'); if (ac) ac.textContent = n.toLocaleString('es-MX');
  }

  function recargar() { cargarKpis(); cargarTabla(); actualizarContadorNav(); }

  // Trae todos los registros que cumplen los filtros, de 1000 en 1000
  async function traerTodo(construir) {
    var todo = [], desde = 0, lote = 1000;
    for (;;) {
      var r = await construir().range(desde, desde + lote - 1);
      if (r.error) throw r.error;
      todo = todo.concat(r.data);
      if (r.data.length < lote) break;
      desde += lote;
    }
    return todo;
  }

  async function exportarConsulta() {
    var btn = $('#btn-export');
    cargando(btn, true, 'Exportando…');
    try {
      var datos = await traerTodo(function () {
        return aplicarFiltros(sb.from('v_agregados').select('*')).order(Q.orden, { ascending: Q.asc, nullsFirst: false }).order('id', { ascending: false });
      });
      var wb = window.ExcelRespaldo.consulta(window.XLSX, datos);
      var nombre = 'CONSULTA_AGREGADOS_' + hoyISO() + '.xlsx';
      window.XLSX.writeFile(wb, nombre);
      toast('ok', 'Excel descargado', datos.length.toLocaleString('es-MX') + ' registros · ' + nombre);
    } catch (e) {
      toast('bad', 'No se pudo exportar', mensajeError(e));
    }
    cargando(btn, false);
  }

  /* =========================================================================
     7. ACCIONES POR REGISTRO
     ========================================================================= */
  var actual = null; // fila sobre la que se actúa
  var fCarga = null;

  function cerrarDialogos() { $$('dialog[open]').forEach(function (d) { d.close(); }); }
  $$('dialog').forEach(function (d) {
    d.addEventListener('click', function (e) { if (e.target.closest('[data-close]')) d.close(); });
    d.addEventListener('cancel', function (e) { if (popAbierto) { e.preventDefault(); cerrarPop(); } });
    d.addEventListener('close', cerrarPop);
  });

  function resumenFila(r) { return 'NUR ' + r.nur + ' · ' + r.cpt + ' · ' + r.proveedor; }

  async function accionFila(e) {
    var b = e.target.closest('button[data-a]'); if (!b) return;
    var r = filas[+b.closest('tr').dataset.i]; if (!r) return;
    actual = r;
    var a = b.dataset.a;
    if (a === 'editar') abrirEdicion(r);
    else if (a === 'carga') abrirCarga(r);
    else if (a === 'baja') abrirBaja(r);
    else if (a === 'hist') abrirHistorial(r);
    else if (a === 'reactivar') reactivar(r);
    else if (a === 'copiar') copiarParaCorreo([r]);
  }

  async function abrirEdicion(r) {
    // Se lee de nuevo para editar la versión más reciente
    var f = await sb.from('v_agregados').select('*').eq('id', r.id).single();
    if (f.error) { toast('bad', 'No se pudo abrir', mensajeError(f.error)); return; }
    var dlg = $('#dlg-edit');
    $('#dlg-edit-sub').textContent = resumenFila(f.data);
    dlg.showModal();
    Formulario($('#form-edit'), {
      modo: 'editar', registro: f.data,
      onGuardado: function () { dlg.close(); cargarTabla(); cargarKpis(); },
      onCancelar: function () { dlg.close(); }
    });
    var primero = $('#form-edit .combo-btn'); if (primero) primero.focus();
  }

  function abrirCarga(r) {
    var dlg = $('#dlg-carga');
    $('#dlg-carga-sub').textContent = resumenFila(r) + ' · solicitado el ' + isoADMY(r.fecha_solicitud);
    $('#carga-error').hidden = true;
    fCarga = DateInput($('[data-date="carga-fecha"]'), { value: hoyISO(), min: function () { return r.fecha_solicitud; }, max: hoyISO, clearable: false, id: 'carga-input' });
    $('[data-date="carga-fecha"]').previousElementSibling.setAttribute('for', 'carga-input');
    dlg.showModal();
    fCarga.focus(); fCarga.input.select();
  }
  $('#carga-save').addEventListener('click', async function () {
    var v = fCarga.get(), err = $('#carga-error'), hoy = hoyISO();
    var msg = !v ? 'Escribe o elige la fecha.' : v === 'invalid' ? 'Fecha no válida. Usa dd/mm/aaaa.' : v > hoy ? 'No puede ser una fecha futura.' : v < actual.fecha_solicitud ? 'No puede ser anterior a la fecha de solicitud (' + isoADMY(actual.fecha_solicitud) + ').' : '';
    if (msg) { err.textContent = msg; err.hidden = false; return; }
    cargando(this, true);
    var r = await sb.from('agregados').update({ fecha_carga: v }).eq('id', actual.id).select('id').single();
    cargando(this, false);
    if (r.error) { err.textContent = mensajeError(r.error); err.hidden = false; return; }
    $('#dlg-carga').close();
    toast('ok', 'Marcado como cargado', 'NUR ' + actual.nur + ' · ' + isoADMY(v));
    cargarTabla(); cargarKpis(); actualizarContadorNav();
  });
  $('#dlg-carga').addEventListener('keydown', function (e) { if (e.key === 'Enter' && !popAbierto) { e.preventDefault(); $('#carga-save').click(); } });

  function abrirBaja(r) {
    $('#dlg-baja-sub').textContent = resumenFila(r);
    $('#baja-motivo').value = '';
    $('#baja-error').hidden = true;
    $('#dlg-baja').showModal();
    $('#baja-motivo').focus();
  }
  $('#baja-save').addEventListener('click', async function () {
    var motivo = limpiar($('#baja-motivo').value);
    if (!motivo) { $('#baja-error').hidden = false; $('#baja-motivo').focus(); return; }
    cargando(this, true, 'Dando de baja…');
    var r = await sb.from('agregados').update({ eliminado: true, motivo_baja: motivo, eliminado_en: new Date().toISOString(), eliminado_por: state.user.email }).eq('id', actual.id).select('id').single();
    cargando(this, false);
    if (r.error) { $('#baja-error').textContent = mensajeError(r.error); $('#baja-error').hidden = false; return; }
    $('#dlg-baja').close();
    toast('ok', 'Registro dado de baja', 'Puedes verlo con “Mostrar bajas” y reactivarlo.');
    cargarTabla(); cargarKpis(); actualizarContadorNav();
  });

  async function reactivar(r) {
    var ok = await confirmar('¿Reactivar este registro?', '<p class="card-hint">' + esc(resumenFila(r)) + '</p><p class="card-hint mt-8">Motivo de la baja: ' + esc(r.motivo_baja || '—') + '</p>', 'Reactivar');
    if (!ok) return;
    var u = await sb.from('agregados').update({ eliminado: false, eliminado_en: null, eliminado_por: null, motivo_baja: null }).eq('id', r.id).select('id').single();
    if (u.error) { toast('bad', 'No se pudo reactivar', mensajeError(u.error)); return; }
    toast('ok', 'Registro reactivado', 'NUR ' + r.nur);
    cargarTabla(); cargarKpis(); actualizarContadorNav();
  }

  var CAMPOS_HIST = {
    anio_surtido: 'Año surtido', cliente_id: 'Cliente', proveedor_id: 'Proveedor', contrarrecibo: 'Contrarrecibo', nur: 'NUR', cpt: 'CPT',
    descripcion: 'Descripción', tipo_servicio_id: 'Tipo de servicio', solicitante_id: 'Solicitante', fecha_solicitud: 'Fecha solicitud',
    fecha_carga: 'Fecha de carga', estatus: 'Estatus', notas: 'Notas', requiere_revision: 'Requiere revisión', eliminado: 'Dado de baja', motivo_baja: 'Motivo de baja'
  };
  var CAT_DE = { cliente_id: 'clientes', proveedor_id: 'proveedores', tipo_servicio_id: 'tipos_servicio', solicitante_id: 'solicitantes' };
  function valorHist(k, v) {
    if (v == null || v === '') return '—';
    if (CAT_DE[k]) return nombreDe(CAT_DE[k], v) || ('#' + v);
    if (k === 'fecha_solicitud' || k === 'fecha_carga') return isoADMY(v);
    if (typeof v === 'boolean') return v ? 'Sí' : 'No';
    return String(v);
  }
  async function abrirHistorial(r) {
    var dlg = $('#dlg-hist'), ul = $('#hist-list');
    $('#dlg-hist-sub').textContent = resumenFila(r);
    ul.innerHTML = '<li class="muted">Cargando…</li>';
    dlg.showModal();
    var h = await sb.from('agregados_historial').select('*').eq('agregado_id', r.id).order('fecha', { ascending: false });
    if (h.error) { ul.innerHTML = '<li>' + esc(mensajeError(h.error)) + '</li>'; return; }
    var items = h.data.map(function (x) {
      var op = { INSERT: 'Alta', UPDATE: 'Cambio', BAJA: 'Baja', DELETE: 'Borrado' }[x.operacion] || x.operacion;
      var filasDiff = '';
      if (x.operacion !== 'INSERT' && x.datos_antes && x.datos_despues) {
        Object.keys(CAMPOS_HIST).forEach(function (k) {
          var a = x.datos_antes[k], d = x.datos_despues[k];
          if (JSON.stringify(a) !== JSON.stringify(d)) filasDiff += '<tr><td>' + CAMPOS_HIST[k] + '</td><td><span class="old">' + esc(valorHist(k, a)) + '</span> → <span class="new">' + esc(valorHist(k, d)) + '</span></td></tr>';
        });
      }
      if (x.operacion === 'INSERT') filasDiff = '<tr><td>Capturado</td><td>' + esc(valorHist('cpt', x.datos_despues && x.datos_despues.cpt)) + ' · ' + esc(valorHist('descripcion', x.datos_despues && x.datos_despues.descripcion)) + '</td></tr>';
      return '<li><div class="h-top"><span class="h-op">' + op + '</span><span>' + esc(fechaHora(x.fecha)) + ' · ' + esc(x.usuario || 'sistema') + '</span></div>' +
        (filasDiff ? '<table>' + filasDiff + '</table>' : '<p class="muted mt-8" style="font-size:12.5px">Sin cambios en los datos visibles.</p>') + '</li>';
    });
    if (r.origen === 'EXCEL_2026') {
      items.push('<li><div class="h-top"><span class="h-op">Migrado del Excel</span><span>Fila ' + r.fila_excel + ' de AGREGADOS 2026</span></div>' +
        '<p class="muted mt-8" style="font-size:12.5px">La copia exacta de la fila original se conserva en la base (tabla agregados_excel_original).</p></li>');
    }
    ul.innerHTML = items.join('') || '<li class="muted">Sin movimientos registrados.</li>';
  }

  /* =========================================================================
     8. CATÁLOGOS (pantalla)
     ========================================================================= */
  var catSel = 'clientes';
  function iniciarCatalogos() {
    pintarLadoCat();
    document.addEventListener('catalogos', function () { pintarLadoCat(); pintarCat(); });
    $('#cat-side').addEventListener('click', function (e) {
      var b = e.target.closest('button[data-cat]'); if (!b) return;
      catSel = b.dataset.cat; $('#cat-q').value = ''; $('#cat-dup').hidden = true; pintarLadoCat(); pintarCat();
    });
    $('#cat-q').addEventListener('input', pintarCat);
    $('#cat-new').addEventListener('input', function () {
      var v = limpiar(this.value), box = $('#cat-dup');
      var par = v ? parecidos(catSel, v) : [];
      box.hidden = !par.length;
      if (par.length) box.innerHTML = icon('i-alert') + '<span>Parecidos ya registrados: <b>' + par.slice(0, 6).map(function (x) { return esc(x.nombre) + (x.activo ? '' : ' (inactivo)'); }).join('</b>, <b>') + '</b></span>';
    });
    $('#cat-add').addEventListener('submit', async function (e) {
      e.preventDefault();
      var inp = $('#cat-new'), v = limpiar(inp.value).toUpperCase();
      if (!v) { inp.focus(); return; }
      var igual = state.cat[catSel].filter(function (x) { return norm(x.nombre) === norm(v); })[0];
      if (igual && igual.activo) { toast('info', 'Ya existe', igual.nombre); return; }
      var par = parecidos(catSel, v).filter(function (x) { return norm(x.nombre) !== norm(v); });
      if (par.length) {
        var ok = await confirmar('¿Dar de alta “' + v + '”?', '<div class="notice warn">' + icon('i-alert') + '<span>Ya existen nombres parecidos: <b>' + par.slice(0, 5).map(function (x) { return esc(x.nombre); }).join('</b>, <b>') + '</b>. Si es el mismo, no lo dupliques.</span></div>', 'Dar de alta');
        if (!ok) return;
      }
      try { await altaCatalogo(catSel, v); inp.value = ''; $('#cat-dup').hidden = true; inp.focus(); }
      catch (err) { toast('bad', 'No se pudo dar de alta', mensajeError(err)); }
    });
    $('#cat-rows').addEventListener('change', async function (e) {
      var sw = e.target.closest('input[data-activo]'); if (!sw) return;
      var id = +sw.dataset.activo;
      var r = await sb.from(catSel).update({ activo: sw.checked }).eq('id', id);
      if (r.error) { sw.checked = !sw.checked; toast('bad', 'No se pudo cambiar', mensajeError(r.error)); return; }
      toast('ok', sw.checked ? 'Activado' : 'Desactivado', nombreDe(catSel, id) + (sw.checked ? '' : ' ya no aparecerá en las listas'));
      await cargarCatalogos();
    });
    $('#cat-rows').addEventListener('click', function (e) {
      var b = e.target.closest('button[data-renombrar]'); if (!b) return;
      var li = b.closest('li'), id = +b.dataset.renombrar, item = state.byId[catSel][id];
      var nameBox = $('.name', li);
      nameBox.innerHTML = '<input class="field" value="' + esc(item.nombre) + '" aria-label="Nuevo nombre" />';
      var inp = $('input', nameBox); inp.focus(); inp.select();
      var hecho = false;
      async function guardar() {
        if (hecho) return; hecho = true;
        var v = limpiar(inp.value).toUpperCase();
        if (!v || v === item.nombre) { pintarCat(); return; }
        var ok = await confirmar('¿Cambiar el nombre?', '<p class="card-hint">“' + esc(item.nombre) + '” → “<b>' + esc(v) + '</b>”</p><p class="card-hint mt-8">El cambio se verá en todos los registros que lo usan.</p>', 'Cambiar nombre');
        if (!ok) { pintarCat(); return; }
        var r = await sb.from(catSel).update({ nombre: v }).eq('id', id);
        if (r.error) toast('bad', 'No se pudo renombrar', mensajeError(r.error));
        else { toast('ok', 'Nombre actualizado', v); state.consultaSucia = true; }
        await cargarCatalogos();
      }
      inp.addEventListener('keydown', function (ev) { if (ev.key === 'Enter') { ev.preventDefault(); guardar(); } if (ev.key === 'Escape') { hecho = true; pintarCat(); } });
      inp.addEventListener('blur', guardar);
    });
    pintarCat();
  }
  function pintarLadoCat() {
    $('#cat-side').innerHTML = CATALOGOS.map(function (c) {
      var n = state.cat[c.key].filter(function (x) { return x.activo; }).length;
      return '<button type="button" role="tab" data-cat="' + c.key + '" aria-current="' + (c.key === catSel) + '"><span>' + c.titulo + '</span><span class="n">' + n + '</span></button>';
    }).join('');
  }
  function pintarCat() {
    var c = CATALOGOS.filter(function (x) { return x.key === catSel; })[0];
    $('#cat-title').textContent = c.titulo;
    $('#cat-new').placeholder = 'Nuevo ' + c.singular;
    var q = norm($('#cat-q').value);
    var lista = state.cat[catSel].filter(function (x) { return !q || norm(x.nombre).indexOf(q) >= 0; })
      .sort(function (a, b) { return (b.activo - a.activo) || a.nombre.localeCompare(b.nombre); });
    $('#cat-rows').innerHTML = lista.length ? lista.map(function (x) {
      return '<li class="' + (x.activo ? '' : 'inactive') + '"><span class="name">' + esc(x.nombre) + '</span>' +
        '<button type="button" class="icon-btn sm plain" data-renombrar="' + x.id + '" title="Cambiar nombre" aria-label="Cambiar nombre de ' + esc(x.nombre) + '">' + icon('i-edit') + '</button>' +
        '<label class="switch" title="' + (x.activo ? 'Activo' : 'Inactivo') + '"><input type="checkbox" data-activo="' + x.id + '"' + (x.activo ? ' checked' : '') + ' aria-label="Activo: ' + esc(x.nombre) + '"><span></span></label></li>';
    }).join('') : '<li class="muted">Sin resultados.</li>';
  }

  /* =========================================================================
     9. CORREOS Y RESPALDO
     ========================================================================= */
  var destinatarios = [], columnasExtra = true;
  var EMAIL_RE = /^[^@\s]+@[^@\s]+\.[^@\s]+$/;
  function iniciarCorreos() {
    $('#mail-add').addEventListener('submit', async function (e) {
      e.preventDefault();
      var nombre = limpiar($('#mail-name').value), email = limpiar($('#mail-email').value).toLowerCase(), box = $('#mail-error');
      box.hidden = true;
      var msg = !nombre ? 'Escribe el nombre.' : !email ? 'Escribe el correo.' : !EMAIL_RE.test(email) ? 'El correo no es válido.' :
        destinatarios.some(function (d) { return d.email.toLowerCase() === email; }) ? 'Ese correo ya está en la lista.' : '';
      if (msg) { box.innerHTML = icon('i-alert') + '<span>' + msg + '</span>'; box.hidden = false; return; }
      var r = await sb.from('destinatarios_alerta').insert({ nombre: nombre.toUpperCase(), email: email, activo: true });
      if (r.error) { box.innerHTML = icon('i-alert') + '<span>' + esc(mensajeError(r.error)) + '</span>'; box.hidden = false; return; }
      $('#mail-name').value = ''; $('#mail-email').value = '';
      toast('ok', 'Destinatario agregado', email);
      cargarDestinatarios();
    });
    $('#mail-rows').addEventListener('change', async function (e) {
      var sw = e.target.closest('input[data-campo]'); if (!sw) return;
      var cambio = {}; cambio[sw.dataset.campo] = sw.checked;
      var r = await sb.from('destinatarios_alerta').update(cambio).eq('id', +sw.dataset.id);
      if (r.error) { sw.checked = !sw.checked; toast('bad', 'No se pudo cambiar', mensajeError(r.error)); return; }
      cargarDestinatarios();
    });
    $('#mail-rows').addEventListener('click', async function (e) {
      var b = e.target.closest('button[data-quitar]'); if (!b) return;
      var d = destinatarios.filter(function (x) { return x.id === +b.dataset.quitar; })[0];
      var ok = await confirmar('¿Quitar destinatario?', '<p class="card-hint">' + esc(d.nombre) + ' — ' + esc(d.email) + ' ya no recibirá correos.</p>', 'Quitar', true);
      if (!ok) return;
      var r = await sb.from('destinatarios_alerta').delete().eq('id', d.id).select('id');
      if (r.error || !r.data || !r.data.length) {
        // Si no hay permiso de borrar (falta correr 02_ajustes_sistema.sql), se desactiva
        r = await sb.from('destinatarios_alerta').update({ activo: false }).eq('id', d.id);
      }
      if (r.error) toast('bad', 'No se pudo quitar', mensajeError(r.error));
      cargarDestinatarios();
    });
    $('#btn-backup').addEventListener('click', descargarRespaldo);
    $('#alert-see').addEventListener('click', function () {
      setTimeout(function () {
        Q.anio = null; pintarFiltroAnios(); Q.ops = OPS.CAMBIO; $('#f-ops').value = Q.ops; $('#filters-more').hidden = false;
        pintarKpiActivo(); recargar();
      }, 0);
    });
  }
  async function cargarDestinatarios() {
    var r = await sb.from('destinatarios_alerta').select('*').order('id');
    var box = $('#mail-rows');
    if (r.error) { box.innerHTML = '<div class="card-pad muted">' + esc(mensajeError(r.error)) + '</div>'; return; }
    destinatarios = r.data;
    columnasExtra = destinatarios.length ? ('recibe_alerta' in destinatarios[0]) : columnasExtra;
    var aviso = !columnasExtra ? '<div class="card-pad" style="padding-top:0"><div class="notice warn">' + icon('i-alert') + '<span>Para elegir quién recibe la alerta y quién el respaldo, corre <b>sql/02_ajustes_sistema.sql</b> en Supabase. Mientras tanto, todos los activos reciben ambos correos.</span></div></div>' : '';
    var visibles = destinatarios.filter(function (d) { return d.activo || !columnasExtra; });
    box.innerHTML = aviso + (visibles.length ? visibles.map(function (d) {
      var sw = function (campo, txt) {
        return '<label><span class="switch"><input type="checkbox" data-campo="' + campo + '" data-id="' + d.id + '"' + (d[campo] ? ' checked' : '') + '><span></span></span>' + txt + '</label>';
      };
      return '<div class="mail-row"><div class="who"><b>' + esc(d.nombre) + '</b><span>' + esc(d.email) + '</span>' +
        (/cambiar\.com$/i.test(d.email) ? '<div class="notice warn mt-8" style="padding:6px 10px">' + icon('i-alert') + '<span>Correo de ejemplo: agrega el real y quita este.</span></div>' : '') + '</div>' +
        (columnasExtra ? sw('recibe_alerta', 'Alerta') + sw('recibe_respaldo', 'Respaldo') : sw('activo', 'Activo') + '<span></span>') +
        '<button type="button" class="icon-btn sm plain bad" data-quitar="' + d.id + '" title="Quitar" aria-label="Quitar ' + esc(d.email) + '">' + icon('i-trash') + '</button></div>';
    }).join('') : '<div class="empty"><strong>Sin destinatarios</strong>Agrega al menos un correo para recibir la alerta y el respaldo.</div>');
  }

  async function descargarRespaldo() {
    var btn = $('#btn-backup');
    cargando(btn, true, 'Generando respaldo…');
    try {
      var datos = await traerTodo(function () { return sb.from('v_agregados').select('*').order('id'); });
      var cats = await Promise.all(CATALOGOS.map(function (c) { return sb.from(c.key).select('nombre,activo').order('nombre'); }));
      var catalogos = {};
      CATALOGOS.forEach(function (c, i) { if (cats[i].error) throw cats[i].error; catalogos[c.key] = cats[i].data; });
      var wb = window.ExcelRespaldo.libro(window.XLSX, { agregados: datos, catalogos: catalogos, generado: new Date() });
      var nombre = window.ExcelRespaldo.nombreArchivo(new Date());
      window.XLSX.writeFile(wb, nombre);
      toast('ok', 'Respaldo descargado', datos.length.toLocaleString('es-MX') + ' registros · ' + nombre);
    } catch (e) {
      toast('bad', 'No se pudo generar el respaldo', mensajeError(e));
    }
    cargando(btn, false);
  }

  /* ---------- Capturados en esta sesión (para copiarlos juntos al correo) ---------- */
  var capturados = [];
  function iniciarNuevo() {
    var panel = document.createElement('div');
    panel.id = 'sesion'; panel.className = 'sesion'; panel.hidden = true;
    $('#form-nuevo').parentNode.appendChild(panel);
    panel.addEventListener('click', function (e) {
      if (e.target.closest('[data-copiar-sesion]')) copiarParaCorreo(capturados);
      if (e.target.closest('[data-vaciar-sesion]')) { capturados = []; pintarSesion(); }
    });
    Formulario($('#form-nuevo'), { modo: 'nuevo', onCreado: function (r) { capturados.push(r); pintarSesion(); } });
  }
  function pintarSesion() {
    var panel = $('#sesion'); if (!panel) return;
    panel.hidden = !capturados.length;
    if (!capturados.length) { panel.innerHTML = ''; return; }
    panel.innerHTML = '<div class="sesion-head"><b>Capturados en esta sesión: ' + capturados.length + '</b><span class="spacer"></span>' +
      '<button type="button" class="btn-primary btn-sm" data-copiar-sesion>' + icon('i-copy') + 'Copiar tabla para correo</button>' +
      '<button type="button" class="btn-ghost btn-sm" data-vaciar-sesion>Vaciar lista</button></div>' +
      '<div style="overflow-x:auto"><table><tr>' + COPIA_ENC.map(function (h) { return '<th>' + h + '</th>'; }).join('') + '</tr>' +
      capturados.map(function (r) { return '<tr>' + copiaValores(r).map(function (v) { return '<td>' + esc(v) + '</td>'; }).join('') + '</tr>'; }).join('') +
      '</table></div><p class="card-hint mt-8">La lista se vacía al cerrar o recargar la página. Lo guardado ya está en la base.</p>';
  }

  /* =========================================================================
     10. NAVEGACIÓN Y ARRANQUE
     ========================================================================= */
  var VISTAS = ['consulta', 'nuevo', 'catalogos', 'correos'];
  function enrutar() {
    if ($('#app').hidden) return;
    var v = (location.hash || '').replace('#', '');
    if (VISTAS.indexOf(v) < 0) v = 'consulta';
    cerrarPop();
    $$('.view').forEach(function (s) { s.hidden = s.dataset.view !== v; });
    $$('.nav .tab').forEach(function (t) { if (t.dataset.view === v) t.setAttribute('aria-current', 'page'); else t.removeAttribute('aria-current'); });
    if (v === 'consulta') {
      if (!state.iniciado.consulta) { state.iniciado.consulta = true; iniciarConsulta(); recargar(); }
      else if (state.consultaSucia) recargar();
    } else if (v === 'nuevo') {
      if (!state.iniciado.nuevo) { state.iniciado.nuevo = true; iniciarNuevo(); }
      setTimeout(function () { var b = $('#form-nuevo .combo-btn'); if (b) b.focus(); }, 30);
    } else if (v === 'catalogos') {
      if (!state.iniciado.catalogos) { state.iniciado.catalogos = true; iniciarCatalogos(); }
    } else if (v === 'correos') {
      if (!state.iniciado.correos) { state.iniciado.correos = true; iniciarCorreos(); }
      cargarDestinatarios(); actualizarContadorNav();
    }
    document.title = { consulta: 'Consulta', nuevo: 'Nuevo agregado', catalogos: 'Catálogos', correos: 'Correos y respaldo' }[v] + ' · Seguimiento de Agregados';
  }

  function iniciarTema() {
    var b = $('#theme-toggle');
    function pintar() {
      var oscuro = document.documentElement.getAttribute('data-theme') === 'dark';
      b.innerHTML = icon(oscuro ? 'i-sun' : 'i-moon');
      b.setAttribute('aria-label', oscuro ? 'Cambiar a modo claro' : 'Cambiar a modo oscuro');
    }
    b.addEventListener('click', function () {
      var t = document.documentElement.getAttribute('data-theme') === 'dark' ? 'light' : 'dark';
      document.documentElement.setAttribute('data-theme', t);
      document.documentElement.classList.toggle('dark', t === 'dark');
      try { localStorage.setItem('agr-theme', t); } catch (e) {}
      pintar();
    });
    pintar();
  }

  async function arrancar() {
    iniciarTema();
    if (!configurado()) { mostrar('setup'); return; }
    sb = window.supabase.createClient(CFG.SUPABASE_URL, CFG.SUPABASE_ANON_KEY, {
      auth: { persistSession: true, autoRefreshToken: true, storageKey: 'agr-auth' }
    });
    iniciarSesionUI();
    window.addEventListener('hashchange', enrutar);
    var s = await sb.auth.getSession();
    if (s.data && s.data.session) await alEntrar(s.data.session);
    else { mostrar('login'); $('#login-email').focus(); }
    sb.auth.onAuthStateChange(function (evento, session) {
      if (evento === 'SIGNED_IN' && session && (!state.user || state.user.id !== session.user.id)) alEntrar(session);
      if (evento === 'SIGNED_OUT') {
        state.user = null; cerrarDialogos(); mostrar('login');
        $('#login-pass').value = ''; $('#login-email').focus();
      }
    });
    // Refresca indicadores cuando vuelves a la pestaña (los días avanzan solos)
    document.addEventListener('visibilitychange', function () {
      if (!document.hidden && state.user && state.iniciado.consulta && !$('#view-consulta').hidden) recargar();
    });
  }

  arrancar();
})();
