# Los Optimistas — despliegue independiente (Google Sheets + Apps Script + Vercel)

Esto reemplaza a Claude artifacts. Ya no depende de cuenta de Claude para nadie del grupo.

## Qué es cada cosa
- **Google Sheets + Apps Script** = la base de datos. Google Sheets guarda los datos en una pestaña "KV"; Apps Script expone una API que el sitio consulta.
- **Google Drive** = donde se guardan las fotos de las tarjetas.
- **GitHub** = donde vive el código.
- **Vercel** = quien construye el sitio a partir del código y lo publica en una URL fija.

## Paso 1 — Google Sheets + Apps Script (backend)
1. Andá a sheets.google.com → hoja nueva. Nombrala "Los Optimistas - DB".
2. Extensiones → Apps Script.
3. Borrá todo el contenido del editor y pegá el archivo `apps-script/Code.gs` de esta carpeta, entero.
4. Íconos de la izquierda → Configuración del proyecto (el engranaje) → "Propiedades del script" → Agregar propiedad:
   - Nombre: `API_TOKEN`
   - Valor: una cadena larga random (podés generarla en uuidgenerator.net). **Guardala, la vas a necesitar en el Paso 3.**
5. Implementar (botón arriba a la derecha) → Nueva implementación → ícono de engranaje → tipo **"Aplicación web"**.
   - Ejecutar como: **tu cuenta**
   - Quién tiene acceso: **Cualquier usuario**
6. Autorizá los permisos que pida (acceso a la hoja y a Drive — son tuyos, es normal que los pida).
7. Te da una URL que termina en `/exec`. **Copiala, la necesitás en el Paso 3.**

Cada vez que yo te dé un `Code.gs` actualizado en el futuro: pegalo reemplazando el anterior, y volvé a "Implementar → Gestionar implementaciones → lápiz de editar → Nueva versión → Implementar". Guardar el archivo solo no alcanza para que el cambio salga a producción.

## Paso 2 — GitHub (dónde vive el código)
1. Creá cuenta en github.com si no tenés.
2. "New repository" → nombre `los-optimistas` → Create.
3. Subí todos los archivos de esta carpeta (excepto `apps-script/`, que se pega directo en Apps Script, no va al repo) usando el botón "uploading an existing file" en la propia web de GitHub — arrastrás la carpeta entera.

## Paso 3 — Vercel (hosting)
1. Creá cuenta en vercel.com, elegí "Continue with GitHub" para que quede conectado directo.
2. "Add New" → "Project" → elegí el repositorio `los-optimistas` → Import.
3. Antes de tocar "Deploy", abrí "Environment Variables" y cargá:
   - `APPS_SCRIPT_URL` = la URL que termina en /exec (SIN el prefijo VITE_ — a propósito, para que quede solo del lado del servidor y nunca aparezca en el código público del sitio)
   - `VITE_API_TOKEN` = el mismo valor que pusiste como API_TOKEN en Apps Script
4. Deploy. Te da una URL fija (tipo `los-optimistas.vercel.app`) — esa es la que le pasás al grupo, para siempre.

## De acá en adelante — cómo se actualiza
1. Te doy el archivo `src/App.jsx` corregido.
2. Lo subís a GitHub reemplazando el que está (en la web de GitHub: abrís el archivo, ícono de lápiz "Edit", pegás el contenido nuevo, "Commit changes").
3. Vercel reconstruye solo en menos de un minuto. Misma URL de siempre.

No hay riesgo de storage separado ni de URLs que cambian — es exactamente el problema que teníamos con Claude artifacts, resuelto de raíz.

## Limitaciones que quedan (léelo antes de cargar datos reales)
- **Apps Script tiene cuotas diarias** de ejecución (cuentas gratuitas de Google). Con el uso de un grupo de 25 amigos jugando ocasionalmente, no debería ser un problema — pero si algún día ves errores intermitentes sin motivo aparente, esta es la primera sospecha.
- **El token de la API queda visible** en el código del sitio publicado (cualquiera que abra las herramientas de desarrollador del navegador lo puede ver). Cumple la misma función que los PIN de la app: desalienta el uso casual, no es seguridad real contra alguien decidido. La URL de Apps Script en sí SÍ queda oculta (vive solo en el servidor de Vercel), pero el token que autoriza las escrituras no.
- **Cada pedido pasa por dos saltos** (sitio → función de Vercel → Apps Script), no uno solo — es la solución al problema de CORS que ya habíamos encontrado en Control DJ, a costa de una fracción de segundo extra por consulta. No debería notarse en el uso real.
- **Las fotos se nombran `fecha_jugador_cancha_id.jpg`** en Drive (acentos y espacios se sacan automáticamente). El `id` al final es un pedazo corto del identificador interno de la tarjeta, para evitar que dos tarjetas del mismo jugador/cancha/fecha choquen de nombre.
- **Los archivos de foto borrados van a la papelera de Drive**, no se eliminan directo — Google los borra automáticamente a los 30 días. Si te preocupa el espacio, podés vaciarla manualmente vos.
