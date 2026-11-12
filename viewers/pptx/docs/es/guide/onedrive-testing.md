---
title: Pruebas de la integración con OneDrive
description: Verifique la conexión compartida de Office con OneDrive mediante transportes locales de prueba y la aceptación con una cuenta autorizada de Microsoft.
---

# Pruebas de la integración con OneDrive

La conexión con OneDrive la comparten Word, Excel, PowerPoint y Visio. Sus transportes automatizados pertenecen solo a los accesorios de prueba. No hay ninguna interfaz de simulación incluida, ningún servicio simulado público ni una demostración de simulación de OneDrive que configurar o desplegar.

Use la [guía de pruebas de Office](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-testing.md) para los comandos, las comprobaciones por formato y la matriz de aceptación con cuentas reales. La [guía de integración](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-integration.md) documenta la configuración de paquetes y frameworks, la autenticación, el uso compartido y la recuperación de conflictos.

Las pruebas de regresión automatizadas pueden verificar la secuencia de solicitudes a Graph y el cableado de guardado del editor sin credenciales de Microsoft. Cubren instantáneas e identidad de archivo, restricciones de permisos, condiciones de revisión, ediciones locales conservadas, copias separadas en la nube y gestión de errores. No pueden demostrar el inicio de sesión de Microsoft, el consentimiento, el acceso condicional, el CORS del origen desplegado ni la fidelidad de PowerPoint nativo.

El Explorador de Microsoft Graph permite comprobaciones manuales de la API usando el registro de aplicación existente de Microsoft. No valida el registro de su propia aplicación ni su flujo de inicio de sesión. Use una presentación desechable y siga la guía de pruebas; no incluya el ID de cliente del Explorador de Microsoft Graph ni credenciales prestadas como solución de autenticación de su aplicación.

Para editar de verdad a través de la suite, configure el proveedor de tokens de Microsoft del host antes de que cargue su entrada y luego use **Abrir desde OneDrive**. Las cuentas personales o de trabajo reales y las políticas de inquilino de SharePoint siguen siendo casos de aceptación obligatorios. Ninguna prueba que no requiera registro es evidencia de coautoría en vivo de Microsoft ni de edición concurrente sin conflictos garantizada.
