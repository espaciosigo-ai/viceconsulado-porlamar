# ERRORES COMUNES Y SOLUCIONES
## No repetir estos errores durante el desarrollo

---

## ERRORES DE GOOGLE APPS SCRIPT

### Error: "No tienes permiso para acceder a esta aplicación"
**Causa:** Al implementar el Apps Script, seleccionaste "Solo yo" en vez de "Cualquier persona".
**Solución:** Ve a Implementar → Administrar implementaciones → edita y cambia acceso a "Cualquier persona". Re-implementa.

### Error: Los datos no llegan al Sheet
**Causa 1:** La URL del Apps Script en el `index.html` todavía dice `TU_URL_DE_APPS_SCRIPT_AQUI`.
**Solución:** Reemplázala con la URL real de la implementación.

**Causa 2:** Implementaste una versión nueva pero la URL cambió.
**Solución:** Cada nueva implementación genera una URL diferente. Usa la más reciente.

**Causa 3:** CORS / el navegador bloquea la petición.
**Solución:** Verifica que el Apps Script esté implementado como Web App con acceso "Cualquier persona". El fetch debe ser POST sin headers de CORS adicionales.

### Error: El correo de confirmación no llega
**Causa:** Gmail tiene límite de 100 correos/día para cuentas gratuitas.
**Solución:** Para testing no es problema. En producción, si envían más de 100 citas diarias (improbable para un viceconsulado), considerar usar la cuenta de la clienta.

### Error: "Calendar not found" en los logs
**Causa:** El nombre del calendario en el código no coincide exactamente con el que creaste.
**Solución:** Verifica que `CONFIG.CALENDAR_NAME` sea exactamente "Citas Viceconsulado" (o el nombre que le pusiste). Es sensible a mayúsculas y tildes.

### Error: Los eventos se crean en horario incorrecto
**Causa:** La zona horaria del calendario o del Apps Script no es la de Venezuela.
**Solución:** En el editor de Apps Script, ve a Configuración del proyecto (ícono de engranaje) y establece la zona horaria a "(GMT-04:00) Caracas". También verifica la zona horaria del Google Calendar.

---

## ERRORES DE GITHUB PAGES

### Error: La página muestra un 404
**Causa 1:** El archivo no se llama exactamente `index.html` (mayúsculas/minúsculas importan).
**Solución:** Asegúrate de que sea `index.html` (todo en minúsculas) en la raíz del repositorio.

**Causa 2:** GitHub Pages no está activado o apunta a otra rama.
**Solución:** Settings → Pages → Source: Deploy from branch → main → / (root) → Save.

**Causa 3:** Acabas de hacer push y no han pasado 2-3 minutos.
**Solución:** Espera. GitHub Pages tarda en construir la primera vez.

### Error: Los cambios no se reflejan en la página
**Causa:** Cache del navegador.
**Solución:** Ctrl+Shift+R para forzar recarga. O abre en pestaña privada/incógnito.

### Error: Los enlaces de formularios del gobierno no abren
**Causa:** Algunos enlaces del Ministerio de Exteriores tienen caracteres especiales (tildes, espacios codificados).
**Solución:** Prueba cada enlace antes de publicar. Si alguno falla, busca la versión actualizada en la web oficial.

---

## ERRORES DE WHATSAPP BUSINESS

### Error: El mensaje de bienvenida no se envía
**Causa:** Solo se envía la primera vez que alguien escribe, o si no ha escrito en 14 días.
**Solución:** Esto es normal. No se reenvía a contactos frecuentes. Para testing, usa un número que no haya escrito antes.

### Error: Las respuestas rápidas no aparecen
**Causa:** Hay que escribir `/` (barra) en el chat para que aparezcan.
**Solución:** Escribir `/pasaporte`, `/cita`, etc. O tocar el ícono de clip → Respuestas rápidas.

### Error: El mensaje de ausencia se envía en horario laboral
**Causa:** El horario configurado en WhatsApp Business no coincide con el real.
**Solución:** Ajustes → Herramientas → Mensaje de ausencia → Programar → "Fuera del horario comercial" y verificar que el horario del perfil sea correcto.

---

## ERRORES DE DISEÑO / HTML

### Error: La página se ve rota en iPhone
**Causa:** Faltan los meta tags de viewport.
**Solución:** Ya está incluido el `<meta name="viewport">` en el HTML. Si sigue habiendo problemas, verificar que no haya CSS con tamaños fijos que superen el ancho de pantalla.

### Error: Los formularios descargables del gobierno dan error
**Causa:** El Ministerio de Exteriores cambia las URLs de sus documentos periódicamente.
**Solución:** Verificar los enlaces cada 3-6 meses contra la web oficial: https://www.exteriores.gob.es/Consulados/caracas/

### Error: El botón de WhatsApp no abre la app
**Causa:** El número en la URL `wa.me/NUMERO` tiene un formato incorrecto.
**Solución:** El número debe ser solo dígitos, sin +, sin espacios, sin guiones. Formato correcto: `wa.me/584121234567`

---

## ERRORES DE FLUJO / LÓGICA DE NEGOCIO

### Error: Se agendaron 2 citas en el mismo horario
**Causa:** Dos personas enviaron el formulario casi al mismo tiempo para la misma fecha.
**Solución:** El Apps Script busca el primer slot disponible, pero si dos requests llegan simultáneamente puede haber race condition. En la práctica es muy improbable con el volumen de un viceconsulado. Si ocurre, la clienta puede mover uno manualmente en Calendar.

### Error: Alguien agendó cita un día feriado
**Causa:** El selector de fecha del HTML solo bloquea fines de semana, no feriados.
**Solución:** Los feriados se manejan manualmente. La clienta debe revisar el Sheet cada mañana y contactar al solicitante si la fecha cae en feriado. En el futuro se puede agregar una lista de feriados al script.

### Error: La clienta no sabe usar el sistema
**Causa:** Falta capacitación.
**Solución:** Entregarle el archivo `docs/GUIA-CLIENTA.md` y hacer una sesión de 15 minutos por videollamada mostrándole el Sheet y el Calendar.

---

## CHECKLIST ANTES DE PASAR A PRODUCCIÓN

- [ ] Todos los `CAMBIAR:` en el HTML han sido reemplazados
- [ ] La URL del Apps Script está actualizada en el HTML
- [ ] El Google Sheet tiene los encabezados (función crearEncabezados ejecutada)
- [ ] El Calendar "Citas Viceconsulado" existe
- [ ] La zona horaria es America/Caracas en Apps Script y Calendar
- [ ] El formulario envía datos correctamente al Sheet
- [ ] El correo de confirmación llega
- [ ] Los eventos se crean en Calendar
- [ ] WhatsApp Business está configurado con perfil, bienvenida y respuestas rápidas
- [ ] Todos los enlaces a formularios del gobierno funcionan
- [ ] La página se ve bien en celular
- [ ] GitHub Pages está activo y la URL funciona
- [ ] La clienta tiene acceso al Sheet y Calendar
