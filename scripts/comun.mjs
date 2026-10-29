/*
  Funciones compartidas por alerta.mjs y respaldo.mjs (corren en GitHub Actions).
  Variables necesarias (GitHub → Settings → Secrets and variables → Actions):
    SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY,
    SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASS, MAIL_FROM
  Opcional (Variables): APP_URL → dirección de la página en GitHub Pages
*/
import { createRequire } from 'node:module';
import { createClient } from '@supabase/supabase-js';
import nodemailer from 'nodemailer';

const require = createRequire(import.meta.url);
export const XLSX = require('xlsx');
require('../js/excel-respaldo.js'); // mismo generador de Excel que usa la página
export const ExcelRespaldo = globalThis.ExcelRespaldo;

export const TZ = 'America/Mexico_City';

export function env(nombre, opcional = false) {
  const v = (process.env[nombre] || '').trim();
  if (!v && !opcional) throw new Error(`Falta la variable ${nombre}. Agrégala en GitHub → Settings → Secrets and variables → Actions.`);
  return v;
}

export const sb = createClient(env('SUPABASE_URL'), env('SUPABASE_SERVICE_ROLE_KEY'), {
  auth: { persistSession: false, autoRefreshToken: false }
});

export function hoyISO() {
  return new Intl.DateTimeFormat('en-CA', { timeZone: TZ, year: 'numeric', month: '2-digit', day: '2-digit' }).format(new Date());
}
export function dmy(iso) {
  if (!iso) return '';
  const [y, m, d] = String(iso).slice(0, 10).split('-');
  return `${d}/${m}/${y}`;
}
export function esc(s) {
  return String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

// Trae todos los registros de 1000 en 1000 (Supabase limita cada consulta)
export async function traerTodo(construir) {
  let todo = [];
  for (let desde = 0; ; desde += 1000) {
    const { data, error } = await construir().range(desde, desde + 999);
    if (error) throw new Error(`Supabase: ${error.message}`);
    todo = todo.concat(data);
    if (data.length < 1000) break;
  }
  return todo;
}

// tipo: 'alerta' | 'respaldo'
export async function destinatarios(tipo) {
  const { data, error } = await sb.from('destinatarios_alerta').select('*').eq('activo', true).order('id');
  if (error) throw new Error(`No se pudieron leer los destinatarios: ${error.message}`);
  const campo = tipo === 'alerta' ? 'recibe_alerta' : 'recibe_respaldo';
  const lista = data.filter((d) => !(campo in d) || d[campo]);
  const reales = lista.filter((d) => !/cambiar\.com$/i.test(d.email));
  if (reales.length < lista.length) console.warn('⚠️  Se omitieron correos de ejemplo (@cambiar.com). Cámbialos en la pantalla "Correos y respaldo".');
  return reales;
}

let transporte = null;
export async function enviarCorreo({ para, asunto, html, texto, adjuntos }) {
  if (!para.length) { console.warn('⚠️  No hay destinatarios activos: no se envió correo.'); return false; }
  if (!transporte) {
    const port = Number(env('SMTP_PORT', true) || 587);
    transporte = nodemailer.createTransport({
      host: env('SMTP_HOST'),
      port,
      secure: port === 465,
      requireTLS: port === 587,
      auth: { user: env('SMTP_USER'), pass: env('SMTP_PASS') }
    });
  }
  const info = await transporte.sendMail({
    from: env('MAIL_FROM', true) || env('SMTP_USER'),
    to: para.map((d) => `"${d.nombre}" <${d.email}>`).join(', '),
    subject: asunto,
    html,
    text: texto,
    attachments: adjuntos || []
  });
  console.log(`✉️  Correo enviado a ${para.map((d) => d.email).join(', ')} (${info.messageId})`);
  return true;
}

// Plantilla de correo con el mismo estilo de la página
export function plantilla(titulo, cuerpo) {
  const app = env('APP_URL', true);
  return `<!doctype html><html><body style="margin:0;background:#f4f5fa;font-family:'Segoe UI',Roboto,Arial,sans-serif;color:#0f172a">
  <div style="max-width:880px;margin:0 auto;padding:24px 16px">
    <div style="display:flex;align-items:center;gap:10px;margin-bottom:16px">
      <div style="width:34px;height:34px;border-radius:10px;background:#7c3aed;color:#fff;font-weight:700;text-align:center;line-height:34px">✓</div>
      <div style="font-weight:700;font-size:16px">Seguimiento de Agregados</div>
    </div>
    <div style="background:#fff;border:1px solid #e2e8f0;border-radius:16px;padding:22px">
      <h1 style="margin:0 0 12px;font-size:19px;letter-spacing:-.01em">${titulo}</h1>
      ${cuerpo}
      ${app ? `<p style="margin:20px 0 0"><a href="${esc(app)}#consulta" style="display:inline-block;background:#7c3aed;color:#fff;text-decoration:none;font-weight:600;padding:10px 16px;border-radius:10px">Abrir el sistema</a></p>` : ''}
    </div>
    <p style="font-size:12px;color:#64748b;margin-top:14px">Correo automático. Los destinatarios se administran en la pantalla “Correos y respaldo”.</p>
  </div></body></html>`;
}
