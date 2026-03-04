# GUÍA DE INSTALACIÓN — Paso a Paso
## Viceconsulado Honorario de España en Nueva Esparta

Todo lo que necesitas hacer, en orden. Copia y pega cada comando.

---

## PASO 1 — Clonar el repositorio en tu PC

Abre la terminal de VS Code (Ctrl + `) y ejecuta:

```bash
git clone https://github.com/espaciosigo-ai/viceconsulado-porlamar.git
cd viceconsulado-porlamar
```

Si ya lo tienes clonado, solo entra a la carpeta:
```bash
cd viceconsulado-porlamar
```

---

## PASO 2 — Copiar los archivos del proyecto

Copia el archivo `index.html` que te di a la carpeta raíz del proyecto.
Copia la carpeta `scripts/` y `docs/` también.

La estructura debe quedar así:
```
viceconsulado-porlamar/
├── index.html          ← La página web
├── scripts/
│   └── apps-script-citas.js  ← Código para Google Sheets
├── docs/
│   ├── GUIA-INSTALACION.md   ← Este archivo
│   ├── GUIA-WHATSAPP.md      ← Cómo configurar WhatsApp Business
│   ├── GUIA-CLIENTA.md       ← Manual para la clienta
│   └── ERRORES-COMUNES.md    ← Errores y soluciones
└── README.md
```

---

## PASO 3 — Configurar Google Sheets

1. Abre el Google Sheet "Citas Viceconsulado" con la cuenta espaciosigo@gmail.com
2. Ve a **Extensiones → Apps Script**
3. Borra todo el código que aparece
4. Copia y pega TODO el contenido del archivo `scripts/apps-script-citas.js`
5. Haz clic en **Guardar** (ícono de diskette o Ctrl+S)
6. En el selector de funciones (arriba), selecciona `crearEncabezados`
7. Haz clic en **Ejecutar** (▶️)
8. Te pedirá permisos — acepta todo (es tu cuenta)
9. Verifica que el Sheet ahora tiene los encabezados con fondo rojo

---

## PASO 4 — Desplegar el Apps Script como Web App

1. En el editor de Apps Script, haz clic en **Implementar → Nueva implementación**
2. En "Tipo", selecciona **App web**
3. Configura:
   - Descripción: "Citas Viceconsulado v1"
   - Ejecutar como: **Yo** (espaciosigo@gmail.com)
   - Quién tiene acceso: **Cualquier persona**
4. Haz clic en **Implementar**
5. Te dará una URL larga que empieza con `https://script.google.com/macros/...`
6. **COPIA ESA URL** — la necesitas para el siguiente paso

---

## PASO 5 — Conectar la página web con Google Sheets

1. Abre el archivo `index.html` en VS Code
2. Busca (Ctrl+F) el texto: `TU_URL_DE_APPS_SCRIPT_AQUI`
3. Reemplázalo con la URL que copiaste en el paso anterior
4. Guarda el archivo

---

## PASO 6 — Personalizar datos del Viceconsulado

En el archivo `index.html`, busca cada `CAMBIAR:` y reemplaza:

| Buscar | Reemplazar con |
|--------|---------------|
| `584121234567` (aparece varias veces) | Número real de WhatsApp sin + ni espacios |
| `+58 XXX-XXXXXXX` | Número de WhatsApp formateado |
| `[Dirección por confirmar]` | Dirección real del viceconsulado |
| `8:00 AM — 1:00 PM` | Horario real (si es diferente) |
| `@[por definir]` | Nombre real de Instagram |
| El `#` en "Abrir en Google Maps" | URL real de Google Maps |

---

## PASO 7 — Subir a GitHub Pages

En la terminal de VS Code:

```bash
git add .
git commit -m "Página del Viceconsulado - versión inicial"
git push origin main
```

Si te pide credenciales de GitHub, ingresa tu usuario y un token de acceso personal.

---

## PASO 8 — Activar GitHub Pages

1. Ve a https://github.com/espaciosigo-ai/viceconsulado-porlamar
2. Haz clic en **Settings** (pestaña superior)
3. En el menú lateral, haz clic en **Pages**
4. En "Source", selecciona **Deploy from a branch**
5. En "Branch", selecciona **main** y carpeta **/ (root)**
6. Haz clic en **Save**
7. Espera 2-3 minutos
8. Tu página estará en: **https://espaciosigo-ai.github.io/viceconsulado-porlamar/**

---

## PASO 9 — Verificar que todo funciona

1. Abre la URL de tu página
2. Navega por las secciones: Inicio, Trámites, Citas, Ubicación
3. Llena el formulario de citas con datos de prueba
4. Verifica que los datos llegaron al Google Sheet
5. Verifica que se creó un evento en Google Calendar
6. Verifica que llegó el correo de confirmación

---

## PASO 10 — Configurar WhatsApp Business

Sigue la guía en `docs/GUIA-WHATSAPP.md`

---

## Cuando la clienta apruebe todo:

1. Cambia el correo en Google Sheets: comparte el Sheet con el correo real de la clienta y transfiérele la propiedad
2. Actualiza la URL del Apps Script si es necesario (re-implementar con la cuenta de la clienta)
3. Actualiza los datos placeholder en `index.html`
4. Haz `git push` de nuevo

---

## ¿Algo no funciona?

Revisa el archivo `docs/ERRORES-COMUNES.md`
