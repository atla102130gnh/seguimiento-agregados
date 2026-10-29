/*
  ALERTA DIARIA — solicitudes con más de 4 días sin fecha de carga
  Corre de lunes a viernes 9:00 am (hora de México) desde GitHub Actions.
  Prueba local: SUPABASE_URL=... SUPABASE_SERVICE_ROLE_KEY=... SMTP_...=... node alerta.mjs
*/
import { sb, traerTodo, destinatarios, enviarCorreo, plantilla, XLSX, ExcelRespaldo, dmy, esc, hoyISO } from './comun.mjs';

const MAX_FILAS_CORREO = 250;

const vencidos = await traerTodo(() => sb.from('v_pendientes_vencidos').select('*').order('fecha_solicitud', { ascending: true }).order('id'));
console.log(`Solicitudes con más de 4 días sin cargar: ${vencidos.length}`);

if (!vencidos.length) {
  console.log('✅ No hay solicitudes vencidas. No se envía correo.');
  process.exit(0);
}

const para = await destinatarios('alerta');

// Resumen por proveedor
const porProv = {};
for (const r of vencidos) porProv[r.proveedor] = (porProv[r.proveedor] || 0) + 1;
const topProv = Object.entries(porProv).sort((a, b) => b[1] - a[1]).slice(0, 8);

const th = 'style="text-align:left;padding:8px 10px;font-size:11px;color:#64748b;border-bottom:1px solid #e2e8f0;white-space:nowrap"';
const td = 'style="padding:8px 10px;font-size:12.5px;border-bottom:1px solid #f1f5f9;vertical-align:top"';
const filas = vencidos.slice(0, MAX_FILAS_CORREO).map((r) => `<tr>
  <td ${td}>${esc(r.cliente)}</td><td ${td}>${esc(r.proveedor)}</td>
  <td ${td} nowrap>${esc(r.nur)}</td><td ${td} nowrap>${esc(r.cpt)}</td><td ${td}>${esc(r.descripcion)}</td>
  <td ${td}>${esc(r.solicitante)}</td><td ${td} nowrap>${dmy(r.fecha_solicitud)}</td>
  <td ${td} nowrap><b style="color:#be123c">${r.dias_transcurridos} días</b></td></tr>`).join('');

const cuerpo = `
  <p style="margin:0 0 14px;font-size:14px;line-height:1.6">Hay <b style="color:#be123c">${vencidos.length}</b> ${vencidos.length === 1 ? 'solicitud' : 'solicitudes'} con más de 4 días sin fecha de carga. <b>Operaciones necesita actualizarlas.</b></p>
  <p style="margin:0 0 6px;font-size:12px;color:#64748b">Proveedores con más pendientes</p>
  <p style="margin:0 0 16px;font-size:13px">${topProv.map(([p, n]) => `${esc(p)}: <b>${n}</b>`).join(' &nbsp;·&nbsp; ')}</p>
  <div style="overflow-x:auto"><table style="border-collapse:collapse;width:100%">
    <thead><tr><th ${th}>CLIENTE</th><th ${th}>PROVEEDOR</th><th ${th}>NUR</th><th ${th}>CPT</th><th ${th}>DESCRIPCIÓN</th><th ${th}>SOLICITANTE</th><th ${th}>SOLICITUD</th><th ${th}>DÍAS</th></tr></thead>
    <tbody>${filas}</tbody></table></div>
  ${vencidos.length > MAX_FILAS_CORREO ? `<p style="font-size:12.5px;color:#64748b">Se muestran las ${MAX_FILAS_CORREO} más antiguas. La lista completa va en el Excel adjunto.</p>` : '<p style="font-size:12.5px;color:#64748b">La lista también va en el Excel adjunto.</p>'}`;

const wb = ExcelRespaldo.consulta(XLSX, vencidos);
const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

await enviarCorreo({
  para,
  asunto: `⚠️ ${vencidos.length} ${vencidos.length === 1 ? 'solicitud de agregados requiere' : 'solicitudes de agregados requieren'} actualización de Operaciones`,
  html: plantilla('Solicitudes pendientes de actualizar', cuerpo),
  texto: `Hay ${vencidos.length} solicitudes con más de 4 días sin fecha de carga. Operaciones necesita actualizarlas. Ver Excel adjunto.`,
  adjuntos: [{ filename: `PENDIENTES_OPERACIONES_${hoyISO()}.xlsx`, content: buffer }]
});
