/*
  RESPALDO SEMANAL — cada viernes 5:00 pm (hora de México) desde GitHub Actions
  1. Descarga TODOS los registros (incluye bajas) y los catálogos.
  2. Genera el Excel con el formato del archivo original.
  3. Lo guarda en el bucket PRIVADO "respaldos" de Supabase Storage.
  4. Lo envía por correo a quienes tienen activado "Respaldo".
*/
import { sb, traerTodo, destinatarios, enviarCorreo, plantilla, XLSX, ExcelRespaldo, hoyISO } from './comun.mjs';

const ahora = new Date();
const agregados = await traerTodo(() => sb.from('v_agregados').select('*').order('id'));
const catalogos = {};
for (const c of ['clientes', 'proveedores', 'solicitantes', 'tipos_servicio']) {
  const { data, error } = await sb.from(c).select('nombre,activo').order('nombre');
  if (error) throw new Error(`No se pudo leer ${c}: ${error.message}`);
  catalogos[c] = data;
}

// Verificación: el respaldo debe traer lo mismo que hay en la base
const { count, error: errCount } = await sb.from('agregados').select('id', { count: 'exact', head: true });
if (errCount) throw new Error(errCount.message);
if (count !== agregados.length) throw new Error(`El respaldo trae ${agregados.length} registros pero la base tiene ${count}. No se envía para no dar un respaldo incompleto.`);
if (!agregados.length) throw new Error('La base no devolvió registros. Revisa la conexión y la llave service_role.');

const wb = ExcelRespaldo.libro(XLSX, { agregados, catalogos, generado: ahora });
const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
const nombre = ExcelRespaldo.nombreArchivo(ahora);
console.log(`📦 ${nombre}: ${agregados.length} registros, ${(buffer.length / 1024).toFixed(0)} KB`);

// Guardar en Storage (privado)
const ruta = `${hoyISO().slice(0, 4)}/${nombre}`;
const { error: errUp } = await sb.storage.from('respaldos').upload(ruta, buffer, {
  contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', upsert: true
});
if (errUp) console.warn(`⚠️  No se guardó en Storage (${errUp.message}). ¿Corriste sql/02_ajustes_sistema.sql? El correo se envía de todas formas.`);
else console.log(`💾 Guardado en Storage: respaldos/${ruta}`);

const activos = agregados.filter((r) => !r.eliminado);
const cargados = activos.filter((r) => r.estatus === 'CARGADO').length;
const vencidos = activos.filter((r) => r.estatus_operaciones === 'CAMBIO DE ESTATUS OPERACIONES').length;
const bajas = agregados.length - activos.length;
const fila = (k, v, color) => `<tr><td style="padding:7px 0;color:#64748b;font-size:13px">${k}</td><td style="padding:7px 0;text-align:right;font-weight:700;font-size:14px;${color ? `color:${color}` : ''}">${v.toLocaleString('es-MX')}</td></tr>`;
const cuerpo = `
  <p style="margin:0 0 14px;font-size:14px;line-height:1.6">Adjunto el respaldo semanal con el mismo formato del Excel original (una hoja por año surtido, más bajas, revisión, catálogos y resumen).</p>
  <table style="width:100%;max-width:360px;border-collapse:collapse">
    ${fila('Registros activos', activos.length)}${fila('Cargados', cargados, '#047857')}${fila('No cargados', activos.length - cargados)}
    ${fila('Más de 4 días sin cargar', vencidos, '#be123c')}${fila('Bajas', bajas)}
  </table>
  <p style="font-size:12.5px;color:#64748b;margin-top:14px">También queda guardado en Supabase → Storage → respaldos/${ruta}</p>`;

const para = await destinatarios('respaldo');
await enviarCorreo({
  para,
  asunto: `📦 Respaldo semanal de agregados — ${hoyISO().split('-').reverse().join('/')}`,
  html: plantilla('Respaldo semanal', cuerpo),
  texto: `Respaldo semanal adjunto: ${activos.length} registros activos, ${bajas} bajas.`,
  adjuntos: [{ filename: nombre, content: buffer }]
});

if (errUp && !para.length) throw new Error('El respaldo no se guardó en Storage ni se envió por correo.');
console.log('✅ Respaldo terminado');
