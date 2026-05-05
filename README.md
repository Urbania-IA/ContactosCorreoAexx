Este script de Google Apps Script automatiza la extracción de contactos desde Gmail y los registra en Google Sheets.

Su objetivo es analizar correos recibidos, detectar posibles datos de contacto en la firma del email y guardarlos de forma estructurada por año.

Funcionalidades principales
Procesa correos recientes de las últimas 24 horas.
Permite procesar todos los correos de un año concreto, actualmente 2025.
Permite ejecutar un procesado histórico completo de la bandeja de entrada.
Extrae automáticamente:
Fecha del correo
Nombre del remitente
Email
Teléfono
Página web
Firma del correo
Asunto
ID interno del mensaje
Agrupa los contactos en pestañas anuales del tipo Contactos-2025, Contactos-2026, etc.
Escribe también los resultados en un Google Sheet general compartido.
Marca los correos procesados con la etiqueta de Gmail contacto-procesado.
Evita duplicados comprobando los IDs de mensajes ya registrados.
Incluye una pestaña de progreso llamada 📊 Progreso.
Permite cancelar manualmente el procesado histórico desde el menú.
Funcionamiento del histórico

El procesado histórico está preparado para grandes volúmenes de correos. Como Google Apps Script tiene límite de tiempo de ejecución, el script procesa los correos en lotes de hasta 100 emails.

Cuando se acerca al límite de tiempo, guarda el punto de avance, escribe los contactos encontrados y programa automáticamente una nueva ejecución para continuar desde donde se quedó.

Menú personalizado

Al abrir el Google Sheet, el script añade un menú llamado:

📬 Procesador de Contactos

Desde este menú se pueden ejecutar las siguientes acciones:

▶️ Procesar últimas 24h
📅 Procesar año completo (2025)
🗂️ Procesar histórico completo
🛑 Cancelar histórico en curso
Hojas generadas

El script crea automáticamente las hojas necesarias:

Contactos-[AÑO]: contactos extraídos del usuario actual.
📊 Progreso: registro del estado del proceso.
En el Sheet general compartido, crea también pestañas Contactos-[AÑO] con una columna adicional de usuario.
Criterio de detección de contactos

El script analiza principalmente la firma del correo y las últimas líneas del mensaje. Asigna una puntuación al correo en función de si encuentra teléfono, web, firma válida y cuerpo de mensaje suficientemente relevante.

Solo se insertan los contactos que superan la puntuación mínima configurada.
