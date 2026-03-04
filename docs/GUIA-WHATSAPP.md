# Guía de Configuración — WhatsApp Business
## Viceconsulado Honorario de España en Nueva Esparta

WhatsApp Business es la app gratuita (NO la API). Se descarga normal desde Play Store o App Store. Esta guía configura todo para que la atención sea rápida sin ser un bot.

---

## 1. Descargar e instalar

- Descargar **WhatsApp Business** (ícono verde con una B)
- Registrar con el número del Viceconsulado
- NO usar el mismo número que WhatsApp personal

---

## 2. Configurar el perfil de empresa

Ve a **Ajustes → Herramientas para la empresa → Perfil de empresa**

| Campo | Valor |
|-------|-------|
| Nombre | Viceconsulado Honorario de España — Nueva Esparta |
| Categoría | Organización gubernamental |
| Descripción | Viceconsulado Honorario de España en Isla de Margarita, Venezuela. Atención de trámites consulares con cita previa. |
| Dirección | [Dirección real] Porlamar, Isla de Margarita |
| Horario | Lunes a Viernes 8:00 AM - 1:00 PM |
| Correo | ch.porlamar@maec.es |
| Sitio web | https://espaciosigo-ai.github.io/viceconsulado-porlamar/ |

Foto de perfil: Escudo de España o logo institucional.

---

## 3. Mensaje de bienvenida automático

Ve a **Ajustes → Herramientas para la empresa → Mensaje de bienvenida**

Activar: ✅ Sí
Destinatarios: Todos

Mensaje:

```
¡Bienvenido/a al Viceconsulado Honorario de España en Nueva Esparta! 🇪🇸

Para una atención más rápida, le recomendamos visitar nuestra página web donde encontrará:

📋 Requisitos de todos los trámites
📅 Formulario para solicitar cita
📍 Ubicación y horarios

🌐 https://espaciosigo-ai.github.io/viceconsulado-porlamar/

Si prefiere, puede escribirnos aquí y le atenderemos en horario de Lunes a Viernes de 8:00 AM a 1:00 PM.

Escriba el NÚMERO de lo que necesita:
1️⃣ Pasaportes
2️⃣ Inscripción de Nacimiento
3️⃣ Inscripción de Matrimonio
4️⃣ Otros trámites
5️⃣ Agendar cita
6️⃣ Ubicación y horario
```

---

## 4. Mensaje de ausencia

Ve a **Ajustes → Herramientas para la empresa → Mensaje de ausencia**

Activar: ✅ Sí
Horario: Fuera del horario de atención
Programar: Fuera del horario comercial

Mensaje:

```
Gracias por comunicarse con el Viceconsulado Honorario de España en Nueva Esparta.

Nuestro horario de atención es de Lunes a Viernes de 8:00 AM a 1:00 PM.

Mientras tanto, puede consultar información sobre trámites y solicitar su cita en nuestra página web:
🌐 https://espaciosigo-ai.github.io/viceconsulado-porlamar/

Le responderemos a la brevedad en nuestro próximo día hábil.
```

---

## 5. Respuestas rápidas (las más importantes)

Ve a **Ajustes → Herramientas para la empresa → Respuestas rápidas**

Crea estas respuestas (se activan escribiendo / en el chat):

### /pasaporte
Atajo: pasaporte

```
🛂 *PASAPORTES — Requisitos:*

• Pasaporte anterior (o del progenitor español si es primera vez)
• DNI en vigor, si posee
• 1 foto 32×26mm color, fondo blanco
• Cédula venezolana (original y copia)
• Pago de tasa consular

⚠️ Mayores de 12 años necesitan toma de huellas (solo en Caracas o en operativos especiales).

📋 Requisitos completos y formularios:
https://espaciosigo-ai.github.io/viceconsulado-porlamar/

📅 ¿Desea agendar cita? Puede hacerlo en la web o responda "cita" por aquí.
```

### /nacimiento
Atajo: nacimiento

```
👶 *INSCRIPCIÓN DE NACIMIENTO — Requisitos:*

• Hoja declaratoria de datos (firmada)
• Acta de nacimiento venezolana apostillada
• Cédula del interesado (original y copia)
• 2 fotos 3×4cm fondo blanco
• Documentos del progenitor español (literal, pasaporte, cédula)

✅ Este trámite es GRATUITO

📋 Requisitos completos y formularios:
https://espaciosigo-ai.github.io/viceconsulado-porlamar/

📅 ¿Desea agendar cita? Puede hacerlo en la web o responda "cita".
```

### /matrimonio
Atajo: matrimonio

```
💍 *INSCRIPCIÓN DE MATRIMONIO — Requisitos:*

• Hoja declaratoria de datos
• Acta de matrimonio venezolana apostillada
• Declaración jurada de estado civil
• Documentos de ambos contrayentes

✅ Trámite GRATUITO. Deben acudir AMBOS contrayentes.

📋 Más info: https://espaciosigo-ai.github.io/viceconsulado-porlamar/
```

### /cita
Atajo: cita

```
📅 *SOLICITAR CITA*

Puede agendar su cita de dos formas:

1️⃣ *Por la web (recomendado):*
https://espaciosigo-ai.github.io/viceconsulado-porlamar/
→ Vaya a "Agendar Cita"

2️⃣ *Por aquí:* Envíenos los siguientes datos:
• Nombre completo
• Cédula o Pasaporte
• Trámite que necesita
• Fecha preferida (Lunes a Viernes)
• Correo electrónico

Le confirmaremos fecha y hora disponible.

⚠️ Recuerde: cada cita es individual y para un solo trámite.
```

### /ubicacion
Atajo: ubicacion

```
📍 *UBICACIÓN Y HORARIO*

🏢 [Dirección por confirmar]
Porlamar, Isla de Margarita
Estado Nueva Esparta

🕐 Lunes a Viernes: 8:00 AM — 1:00 PM
📧 ch.porlamar@maec.es

📌 Google Maps: [enlace por agregar]

⚠️ Todos los trámites requieren cita previa.
```

### /caracas
Atajo: caracas

```
🏛️ Algunos trámites solo se realizan en el *Consulado General en Caracas*:

• Visados (todos los tipos)
• Compulsas de documentos
• Certificado de antecedentes penales
• Legalización y Apostilla

📍 Edif. Bancaracas, Piso 7, La Castellana, Caracas
📞 +58 212 2660333
🆘 Emergencias 24h: +58 424 2090264
🌐 https://www.exteriores.gob.es/Consulados/caracas/
```

### /lmd
Atajo: lmd

```
📜 *LEY DE MEMORIA DEMOCRÁTICA*

⚠️ El plazo para solicitar credenciales venció el 22/10/2025.

Si ya tiene credenciales, puede programar cita presencial para presentar su expediente.

Para consultas: cog.caracas.citaslmd@maec.es

📋 Guía de tramitación:
https://www.exteriores.gob.es/DocumentosAuxiliaresSC/Venezuela/CARACAS%20(C)/2025%2006%2009%20Guía%20aplicación%20LMD%20-%20CG%20Caracas.pdf
```

---

## 6. Catálogo (opcional pero recomendado)

Ve a **Ajustes → Herramientas para la empresa → Catálogo**

Crea estos "productos" (cada trámite como un item):

| Nombre | Precio | Descripción |
|--------|--------|-------------|
| Pasaporte — Solicitud/Renovación | Consultar tasa | Solicitud inicial o renovación de pasaporte español |
| Inscripción de Nacimiento | Gratuito | Registro civil para obtener nacionalidad española |
| Inscripción de Matrimonio | Gratuito | Registro de matrimonio en el consulado |
| Inscripción de Defunción | Gratuito | Registro de defunción de ciudadano español |
| Fe de Vida | Consultar | Certificado de existencia |
| Alta Consular | Consultar | Inscripción como residente en registro consular |

Cada uno con un enlace a la sección correspondiente de la web.

---

## 7. Etiquetas para organizar chats

Crea estas etiquetas de colores para clasificar conversaciones:

| Etiqueta | Color | Uso |
|----------|-------|-----|
| Cita pendiente | 🟡 Amarillo | Persona que solicitó cita pero no se ha confirmado |
| Cita confirmada | 🟢 Verde | Cita ya agendada y confirmada |
| Esperando documentos | 🟠 Naranja | Le falta documentación |
| Atendido | 🔵 Azul | Ya fue atendido en el viceconsulado |
| Redirigir a Caracas | 🔴 Rojo | Trámite que no se hace aquí |

---

## Listo

Con esta configuración, el WhatsApp Business del Viceconsulado:
- Saluda automáticamente a quien escriba por primera vez
- Responde fuera de horario con mensaje de ausencia
- Permite responder en 2 toques con las respuestas rápidas
- Muestra los trámites como catálogo
- Siempre redirige a la web para info completa y citas
