---
title: Mapa de la documentación
description: Un mapa de la documentación de pptx-viewer - qué sección leer según si integra un visor, manipula archivos .pptx por código, usa el editor o automatiza con agentes.
---

# Mapa de la documentación

La documentación está organizada según lo que quiera hacer. Use esta página para encontrar la sección correcta y siga el orden de lectura sugerido.

## Elija un camino

| Quiere...                                                                     | Empiece por                                                                                                                                                                                                           | Después                                                                                                                 |
| ----------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------- |
| Integrar un visor o editor de PowerPoint en una aplicación web                | La página de inicio de su framework: [React](/react/getting-started), [Vue 3](/vue/getting-started), [Angular](/angular/getting-started), [Svelte 5](/svelte/getting-started), [Vanilla JS](/vanilla/getting-started) | La referencia de props/API de esa misma sección, luego [Temas](/guide/theming) y [Localización](/es/guide/localization) |
| Leer, editar o generar archivos `.pptx` por código                            | [Core: carga y análisis](/core/loading) y [la API Builder](/core/builder)                                                                                                                                             | [Edición](/core/editing), [Guardado](/core/saving), [el modelo de datos](/es/guide/data-model)                          |
| Convertir presentaciones a Markdown, imágenes, PDF o SVG                      | [Conversor Markdown](/core/converter) y [Exportación SVG](/core/svg-export)                                                                                                                                           | La página Export de su enlace para los formatos del navegador                                                           |
| Ocultar, bloquear o reasignar partes de la interfaz del visor desde el código | [Personalización de la interfaz](/guide/customization) (en inglés)                                                                                                                                                    | Recetas de esa página: modo quiosco, cinta reducida, ajustes fijados                                                    |
| Usar el editor (sin escribir código)                                          | [Guía de usuario](/es/user/)                                                                                                                                                                                          | [Edición de diapositivas](/es/user/editing), [Atajos de teclado](/es/user/shortcuts)                                    |
| Dejar que agentes de IA o scripts trabajen con presentaciones                 | [MCP y herramientas](/packages/mcp)                                                                                                                                                                                   | [CLI del core](/core/cli)                                                                                               |
| Entender el funcionamiento interno de la biblioteca                           | [Arquitectura](/es/guide/architecture)                                                                                                                                                                                | [Conformidad OOXML](/architecture/openxml-conformance), [Limitaciones](/es/guide/limitations)                           |

## Nuevo en el proyecto

Si ningún camino encaja todavía, lea estas tres páginas en orden:

1. [¿Qué es pptx-viewer?](/es/guide/introduction): la familia de paquetes y qué hace cada parte.
2. [Instalación](/es/guide/installation): qué paquete instalar para su stack.
3. [Inicio rápido](/es/guide/quick-start): cuatro flujos completos con la API pública.

## Las secciones de un vistazo

| Sección                                                                                                                | Público                           | Contenido                                                                                                |
| ---------------------------------------------------------------------------------------------------------------------- | --------------------------------- | -------------------------------------------------------------------------------------------------------- |
| **Guía del desarrollador** (esta sección)                                                                              | Desarrolladores                   | Conceptos que se aplican a todos los paquetes: arquitectura, modelo de datos, temas, i18n, limitaciones. |
| **[Core](/core/)**                                                                                                     | Desarrolladores                   | El motor sin interfaz: cargar, editar, construir, guardar, convertir, cifrar. Sin UI.                    |
| **[React](/react/)** / **[Vue](/vue/)** / **[Angular](/angular/)** / **[Svelte](/svelte/)** / **[Vanilla](/vanilla/)** | Desarrolladores                   | Una sección por enlace de interfaz: primeros pasos, props/API, temas, exportación, colaboración.         |
| **[Guía de usuario](/es/user/)**                                                                                       | Usuarios finales del editor       | Cómo ver, editar, presentar, exportar y colaborar en la interfaz del editor. Sin código.                 |
| **[MCP y herramientas](/packages/mcp)**                                                                                | Autores de agentes/automatización | El servidor MCP, sus 73 herramientas y la CLI.                                                           |
| **[Versiones](/releases/)**                                                                                            | Todos                             | Notas de versión por paquete.                                                                            |

::: info Idioma
Las páginas conceptuales de la guía están traducidas al español; las referencias detalladas de los paquetes (Core, React, Vue, Angular, Svelte, Vanilla) están en inglés.
:::
