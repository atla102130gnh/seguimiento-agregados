# Seguimiento de Carta de Agregados

Sistema web para dar de alta, consultar, modificar y dar de baja servicios agregados. Sustituye la hoja **AGREGADOS 2026** del Excel.

- **Base de datos:** Supabase (ya migrada, 526 registros).
- **Página:** GitHub Pages.
- **Alerta diaria y respaldo de los viernes:** GitHub Actions.

---

## Antes de empezar

Ya debes tener hecho esto:

- [x] Proyecto creado en Supabase y migración `migracion_agregados_2026.sql` corrida y verificada.
- [x] Cuenta de GitHub.

Lo que vas a necesitar a la mano:

- La **Project URL**, la **anon / publishable key** y la **service_role / secret key** de Supabase (Project Settings → API).
- Los datos de la cuenta de correo que enviará las alertas (servidor SMTP, usuario y contraseña).

---

## Paso 1. Ajuste en la base de datos (una sola vez)

En Supabase → **SQL Editor** → New query, pega el contenido de **`02_ajustes_sistema.sql`** (viene en la carpeta `sql-base-de-datos`, fuera del repositorio) y dale **Run**.

Esto permite elegir quién recibe la alerta y quién el respaldo, y crea un almacenamiento **privado** llamado `respaldos` para guardar el Excel de cada viernes. No modifica ningún registro.

## Paso 2. Usuarios del sistema

1. Supabase → **Authentication → Users → Add user → Create new user**. Crea a cada persona con su correo y contraseña. Marca *Auto Confirm User* si aparece.
2. En la configuración de inicio de sesión (Sign In / Providers), **desactiva “Allow new users to sign up”**. Así nadie de fuera puede crearse una cuenta.

## Paso 3. Subir los archivos a GitHub

1. Crea un repositorio nuevo, por ejemplo `seguimiento-agregados`, **Public** (GitHub Pages gratuito requiere repositorio público).
2. En el repositorio: **Add file → Upload files**. Arrastra **todo el contenido** de la carpeta `seguimiento-agregados` (incluida la carpeta `.github`, ver nota abajo).
3. **Commit changes**.

> **Nota sobre la carpeta `.github`:** algunas computadoras ocultan las carpetas que empiezan con punto. Si no la ves, actívalo en el explorador de archivos (Windows: Vista → Mostrar → Elementos ocultos; Mac: Cmd + Shift + punto). Sin esa carpeta no funcionan la alerta ni el respaldo automáticos.

> **El repositorio es público, pero tus datos no.** El código no lleva ninguna llave secreta, la base pide iniciar sesión para entregar cualquier dato, y los respaldos no se guardan en el repositorio. **Nunca subas** a GitHub el archivo de migración ni ningún Excel con datos.

## Paso 4. Conectar la página con Supabase

1. En el repositorio, abre **`config.js`** y da clic en el lápiz (Edit).
2. Reemplaza los dos valores:

```js
window.APP_CONFIG = {
  SUPABASE_URL: 'https://xxxxxxxx.supabase.co',     // Project URL
  SUPABASE_ANON_KEY: 'eyJhbGciOi...'                // anon / publishable key
};
```

3. **Commit changes**.

⚠️ Aquí va la llave **anon / publishable**, nunca la **service_role / secret**.

## Paso 5. Activar GitHub Pages

1. Repositorio → **Settings → Pages**.
2. En *Source* elige **Deploy from a branch**, rama **main**, carpeta **/ (root)** → **Save**.
3. En 1 o 2 minutos aparece la dirección, por ejemplo `https://tuusuario.github.io/seguimiento-agregados/`.
4. Copia esa dirección en Supabase → **Authentication → URL Configuration → Site URL**.

Abre la dirección, inicia sesión con uno de los usuarios del paso 2 y listo.

## Paso 6. Correos automáticos (alerta y respaldo)

### 6.1 Guardar los datos secretos

Repositorio → **Settings → Secrets and variables → Actions → New repository secret**. Crea uno por uno:

| Nombre | Qué poner |
|---|---|
| `SUPABASE_URL` | La Project URL |
| `SUPABASE_SERVICE_ROLE_KEY` | La **service_role / secret** key (solo aquí, nunca en la página) |
| `SMTP_HOST` | Servidor de correo. Microsoft 365 / Outlook: `smtp.office365.com` |
| `SMTP_PORT` | `587` |
| `SMTP_USER` | La cuenta que envía, por ejemplo `notificaciones@tuempresa.com` |
| `SMTP_PASS` | Su contraseña |
| `MAIL_FROM` | Remitente que verán, por ejemplo `Seguimiento de Agregados <notificaciones@tuempresa.com>` |

Opcional: en la pestaña **Variables** de esa misma pantalla, crea `APP_URL` con la dirección de la página. Así el correo trae un botón para abrir el sistema.

### 6.2 Elegir quién recibe los correos

Entra al sistema → **Correos y respaldo**. Agrega a las personas y marca para cada una **Alerta**, **Respaldo** o ambos. Quita los dos correos de ejemplo (`@cambiar.com`); el sistema nunca les envía nada.

### 6.3 Probar

Repositorio → **Actions** → elige *Alerta diaria de solicitudes vencidas* → **Run workflow**. Repite con *Respaldo semanal en Excel*. Si sale palomita verde, revisa la bandeja de entrada. Si sale una X roja, abre la ejecución: el mensaje dice qué falta.

### Si el correo de Microsoft 365 da error de autenticación

Muchas empresas tienen apagado **SMTP AUTH** por seguridad. Pídele a TI que lo active solo para la cuenta que envía (Centro de administración de Microsoft 365 → usuario → Correo → Administrar aplicaciones de correo → *SMTP autenticado*). Si usan verificación en dos pasos, quizá necesiten una contraseña de aplicación.

Alternativas si no se puede: una cuenta de Gmail con *contraseña de aplicación* (`smtp.gmail.com`, puerto `587`) o un servicio como Resend o Brevo; en todos los casos solo cambian los Secrets `SMTP_*`.

---

## Horarios automáticos

| Qué | Cuándo (hora de México) |
|---|---|
| Alerta de solicitudes con más de 4 días sin cargar | Lunes a viernes, 9:00 am |
| Respaldo en Excel | Viernes, 5:00 pm |

GitHub puede retrasar unos minutos las tareas programadas. Además, **GitHub pausa las tareas programadas si el repositorio pasa 60 días sin cambios**; te avisa por correo y se reactivan desde la pestaña Actions con un clic.

## Dónde quedan los respaldos

- En el correo de quienes tienen “Respaldo” activado.
- En Supabase → **Storage → respaldos**, una carpeta por año.
- Cuando quieras: botón **Descargar respaldo ahora** en *Correos y respaldo*.

El Excel trae una hoja por año surtido (AGREGADOS 2026, AGREGADOS 2027…) con las columnas del archivo original, más las hojas BAJAS, REVISION, CATALOGOS y RESUMEN. El NUR y el contrarrecibo van como texto para que Excel no los convierta a notación científica.

## Cambio de año

No hay que hacer nada. El 1 de enero el desplegable **Año surtido** muestra el año nuevo por defecto y sigue ofreciendo los anteriores. Los registros nuevos toman el año en curso, pero se puede elegir otro.

## Cómo protege los datos

- Nada se borra: las bajas solo ocultan el registro y guardan el motivo; se puede reactivar.
- Cada alta, cambio y baja queda en el historial con usuario, fecha y valores anteriores.
- La copia exacta de las 526 filas del Excel original se conserva en la tabla `agregados_excel_original`, que no se puede modificar.

## Estructura del proyecto

```
index.html              Página
config.js               Conexión a Supabase (lo llenas tú)
css/                    Estilos y tipografía
js/app.js               Lógica del sistema
js/excel-respaldo.js    Genera el Excel (lo usan la página y el respaldo automático)
vendor/                 Supabase JS 2.45.4 y SheetJS 0.18.5 (copias locales)
scripts/                Alerta y respaldo (GitHub Actions)
.github/workflows/      Horarios de la alerta y el respaldo
```
