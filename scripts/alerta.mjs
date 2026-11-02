/*
  ALERTAS DIARIAS — lunes a viernes 9:00 am (hora de México), desde GitHub Actions
  Envía dos correos independientes a quienes tienen "Alerta" activada:

    1. VENCIDAS:    solicitudes con MÁS de 4 días sin fecha de carga.
    2. POR VENCER:  solicitudes sin fecha de carga que aún están dentro de los
                    4 días (aviso preventivo: cuáles vencen hoy, mañana, etc.).

  Si alguna lista está vacía, ese correo no se envía.
  Prueba local: SUPABASE_URL=... SUPABASE_SERVICE_ROLE_KEY=... SMTP_...=... node alerta.mjs
*/
import { sb, traerTodo, destinatarios, enviarCorreo, plantilla, XLSX, ExcelRespaldo, dmy, esc, hoyISO } from './comun.mjs';

const DIAS_LIMITE = 4;
const MAX_FILAS_CORREO = 250;
const th = 'style="text-align:left;padding:8px 10px;font-size:11px;color:#64748b;border-bottom:1px solid #e2e8f0;white-space:nowrap"';
const td = 'style="padding:8px 10px;font-size:12.5px;border-bottom:1px solid #f1f5f9;vertical-align:top"';

// Fecha 'AAAA-MM-DD' + n días
function sumarDias(iso, n) {
  const [y, m, d] = iso.split('-').map(Number);
  const f = new Date(Date.UTC(y, m - 1, d + n));
  return f.toISOString().slice(0, 10);
}

const para = await destinatarios('alerta');
let errores = 0;

/* ---------------------------------------------------------------------------
   1. VENCIDAS (más de 4 días)
   --------------------------------------------------------------------------- */
const vencidos = await traerTodo(() => sb.from('v_pendientes_vencidos').select('*').order('fecha_solicitud', { ascending: true }).order('id'));
console.log(`Vencidas (más de ${DIAS_LIMITE} días sin cargar): ${vencidos.length}`);

if (vencidos.length) {
  const porProv = {};
  for (const r of vencidos) porProv[r.proveedor] = (porProv[r.proveedor] || 0) + 1;
  const topProv = Object.entries(porProv).sort((a, b) => b[1] - a[1]).slice(0, 8);

  const filas = vencidos.slice(0, MAX_FILAS_CORREO).map((r) => `<tr>
    <td ${td}>${esc(r.cliente)}</td><td ${td}>${esc(r.proveedor)}</td>
    <td ${td} nowrap>${esc(r.nur)}</td><td ${td} nowrap>${esc(r.cpt)}</td><td ${td}>${esc(r.descripcion)}</td>
    <td ${td}>${esc(r.solicitante)}</td><td ${td} nowrap>${dmy(r.fecha_solicitud)}</td>
    <td ${td} nowrap><b style="color:#be123c">${r.dias_transcurridos} días</b></td></tr>`).join('');

  const cuerpo = `
    <p style="margin:0 0 14px;font-size:14px;line-height:1.6">Hay <b style="color:#be123c">${vencidos.length}</b> ${vencidos.length === 1 ? 'solicitud' : 'solicitudes'} con más de ${DIAS_LIMITE} días sin fecha de carga. <b>Operaciones necesita actualizarlas.</b></p>
    <p style="margin:0 0 6px;font-size:12px;color:#64748b">Proveedores con más pendientes</p>
    <p style="margin:0 0 16px;font-size:13px">${topProv.map(([p, n]) => `${esc(p)}: <b>${n}</b>`).join(' &nbsp;·&nbsp; ')}</p>
    <div style="overflow-x:auto"><table style="border-collapse:collapse;width:100%">
      <thead><tr><th ${th}>CLIENTE</th><th ${th}>PROVEEDOR</th><th ${th}>NUR</th><th ${th}>CPT</th><th ${th}>DESCRIPCIÓN</th><th ${th}>SOLICITANTE</th><th ${th}>SOLICITUD</th><th ${th}>DÍAS</th></tr></thead>
      <tbody>${filas}</tbody></table></div>
    ${vencidos.length > MAX_FILAS_CORREO ? `<p style="font-size:12.5px;color:#64748b">Se muestran las ${MAX_FILAS_CORREO} más antiguas. La lista completa va en el Excel adjunto.</p>` : '<p style="font-size:12.5px;color:#64748b">La lista también va en el Excel adjunto.</p>'}`;

  try {
    const buffer = XLSX.write(ExcelRespaldo.consulta(XLSX, vencidos), { type: 'buffer', bookType: 'xlsx' });
    await enviarCorreo({
      para,
      asunto: `⚠️ ${vencidos.length} ${vencidos.length === 1 ? 'solicitud de agregados requiere' : 'solicitudes de agregados requieren'} actualización de Operaciones`,
      html: plantilla('Solicitudes pendientes de actualizar', cuerpo),
      texto: `Hay ${vencidos.length} solicitudes con más de ${DIAS_LIMITE} días sin fecha de carga. Operaciones necesita actualizarlas. Ver Excel adjunto.`,
      adjuntos: [{ filename: `PENDIENTES_OPERACIONES_${hoyISO()}.xlsx`, content: buffer }]
    });
  } catch (e) { errores++; console.error('❌ No se pudo enviar el correo de vencidas:', e.message); }
} else {
  console.log('✅ No hay solicitudes vencidas: no se envía ese correo.');
}

/* ---------------------------------------------------------------------------
   2. POR VENCER (sin fecha de carga, todavía dentro de los 4 días)
   --------------------------------------------------------------------------- */
const porVencer = await traerTodo(() => sb.from('v_agregados').select('*')
  .eq('eliminado', false)
  .neq('estatus', 'CARGADO')
  .eq('estatus_operaciones', 'EN TIEMPO')
  .order('fecha_solicitud', { ascending: true }).order('id'));
console.log(`Por vencer (dentro de ${DIAS_LIMITE} días, sin cargar): ${porVencer.length}`);

if (porVencer.length) {
  // Días que le quedan: 0 = vence hoy, 1 = vence mañana…
  const conPlazo = porVencer.map((r) => ({ ...r, restan: Math.max(0, DIAS_LIMITE - r.dias_transcurridos), limite: sumarDias(r.fecha_solicitud, DIAS_LIMITE) }));
  const etiqueta = (n) => (n === 0 ? 'Vence hoy' : n === 1 ? 'Vence mañana' : `Vence en ${n} días`);
  const color = (n) => (n === 0 ? '#be123c' : n === 1 ? '#b45309' : '#0f172a');
  const hoy = conPlazo.filter((r) => r.restan === 0).length;
  const manana = conPlazo.filter((r) => r.restan === 1).length;

  const resumen = [hoy && `<b style="color:#be123c">${hoy}</b> ${hoy === 1 ? 'vence' : 'vencen'} hoy`, manana && `<b style="color:#b45309">${manana}</b> mañana`, (conPlazo.length - hoy - manana) && `<b>${conPlazo.length - hoy - manana}</b> en los próximos días`].filter(Boolean).join(', ');

  const filas = conPlazo.map((r) => `<tr>
    <td ${td} nowrap><b style="color:${color(r.restan)}">${etiqueta(r.restan)}</b><div style="font-size:11px;color:#64748b">límite ${dmy(r.limite)}</div></td>
    <td ${td}>${esc(r.proveedor)}</td><td ${td} nowrap>${esc(r.contrarrecibo || '')}</td>
    <td ${td} nowrap>${esc(r.nur)}</td><td ${td} nowrap>${esc(r.cpt)}</td><td ${td}>${esc(r.descripcion)}</td>
    <td ${td}>${esc(r.tipo_servicio)}</td><td ${td} nowrap>${dmy(r.fecha_solicitud)}</td></tr>`).join('');

  const cuerpo = `
    <p style="margin:0 0 14px;font-size:14px;line-height:1.6">Hay <b>${conPlazo.length}</b> ${conPlazo.length === 1 ? 'solicitud' : 'solicitudes'} sin fecha de carga que todavía están en tiempo: ${resumen}. <b>Conviene cargarlas antes de que pasen a vencidas.</b></p>
    <div style="overflow-x:auto"><table style="border-collapse:collapse;width:100%">
      <thead><tr><th ${th}>PLAZO</th><th ${th}>PROVEEDOR</th><th ${th}>CONTRARRECIBO</th><th ${th}>NUR</th><th ${th}>CPT</th><th ${th}>DESCRIPCIÓN</th><th ${th}>TIPO</th><th ${th}>SOLICITUD</th></tr></thead>
      <tbody>${filas}</tbody></table></div>
    <p style="font-size:12.5px;color:#64748b">El plazo es de ${DIAS_LIMITE} días naturales desde la fecha de solicitud. Al capturar la fecha de carga, la solicitud sale de este aviso.</p>`;

  try {
    await enviarCorreo({
      para,
      asunto: `⏳ ${conPlazo.length} ${conPlazo.length === 1 ? 'solicitud por vencer' : 'solicitudes por vencer'}${hoy ? ` (${hoy} ${hoy === 1 ? 'vence' : 'vencen'} hoy)` : ''}`,
      html: plantilla('Solicitudes por vencer', cuerpo),
      texto: `Hay ${conPlazo.length} solicitudes sin fecha de carga por vencer. ${hoy} vencen hoy.`
    });
  } catch (e) { errores++; console.error('❌ No se pudo enviar el correo de por vencer:', e.message); }
} else {
  console.log('✅ No hay solicitudes por vencer: no se envía ese correo.');
}

if (errores) process.exit(1);
