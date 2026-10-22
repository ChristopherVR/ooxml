# Edición de texto

Todos los frameworks montan el mismo editor como componente web. Los comandos, los
atajos de teclado, la búsqueda, los metadatos de idioma y la conversión de documentos
tienen una única implementación.

## Buscar y reemplazar

Use **Buscar y reemplazar** o **Ctrl+F** (**Comando+F** en macOS) mientras el editor
tenga el foco. Escriba una frase de búsqueda literal y use **Buscar siguiente** o
**Buscar anterior** para recorrer las coincidencias. **Coincidir mayúsculas y minúsculas**
controla la distinción entre mayúsculas y minúsculas. Escape cierra el panel y devuelve
el foco al documento.

La búsqueda abarca fragmentos de texto contiguos con distinto formato dentro de un
párrafo. No coincide entre párrafos, celdas de tabla ni saltos de línea forzados. Evita
las coincidencias que dividirían un grafema Unicode, como parte de una secuencia
combinada o de un emoji compuesto. Las búsquedas son literales; no se admiten las
expresiones regulares ni la coincidencia lingüística específica de cada idioma.

**Reemplazar** cambia la coincidencia actual; **Reemplazar todo** cambia todas las
coincidencias del documento en una única operación que se puede deshacer. El texto de
reemplazo toma el formato del primer carácter coincidente. El texto fuera de la
coincidencia conserva su formato. La búsqueda funciona en modo de solo lectura; el
reemplazo está desactivado. Las comprobaciones de preservación de DOCX existentes
siguen aplicándose al exportar ediciones alrededor de contenido de origen no
compatible.

## Documentos multilingües

La dirección del párrafo y el idioma del texto son propiedades independientes. Use
**Dirección del párrafo** para el diseño de izquierda a derecha o de derecha a
izquierda, y **Dirección del fragmento** para una anulación explícita de la dirección
dentro de un párrafo. Los controles de idioma almacenan las etiquetas de idioma normal,
de Asia oriental y de escritura compleja de Word. Los metadatos de idioma no traducen
el texto ni instalan un corrector ortográfico.

El editor conserva el texto Unicode, incluidos árabe, hebreo, CJK, caracteres
combinados y emoji. Las fuentes del navegador, el modelado de glifos y el diseño
bidireccional determinan el aspecto en pantalla. Estos controles no establecen una
tipografía de escritura compleja equivalente a la de Word, ni su paginación, ni el
comportamiento de los IME del sistema operativo.

El recuento de palabras usa la segmentación de palabras Unicode del navegador, incluida
la segmentación por diccionario para escrituras sin espacios cuando está disponible.
Los recuentos pueden diferir entre versiones de ICU y respecto a Microsoft Word. Los
navegadores antiguos usan una alternativa menos precisa.

## Integración de coautoría

La [guía de colaboración](/es/collaboration) describe el protocolo de edición
versionado y cómo una aplicación conecta el editor compartido con su autoridad. Es una
API de integración, no un servicio de colaboración alojado. El control de acceso, la
red, el almacenamiento y la política de reconexión corresponden a la aplicación.

## Idioma de la interfaz

Establezca la propiedad `locale` del editor compartido (o la prop `locale` del
adaptador) en `en`, `fr`, `de`, `es` o `zh-CN`. Las variantes regionales y de escritura
se resuelven al idioma compatible más cercano (`fr-CA`, `de-DE`, `es-MX`, `zh-Hans`,
`zh-SG`); los idiomas no compatibles, incluido el chino tradicional (`zh-TW`,
`zh-Hant`), recurren al inglés. `EDITOR_LOCALES` enumera los códigos y
`normalizeEditorLocale()` convierte cualquier etiqueta en uno de ellos. Cada idioma
define todas las claves de texto (se comprueba en tiempo de compilación y con pruebas).
Las cadenas en alemán, español y chino simplificado no han sido revisadas por hablantes
nativos. Cambiar el idioma de la interfaz conserva el documento actual, la selección, el
historial de deshacer y los metadatos de idioma del documento.

## Estilos de párrafo importados

El selector de estilos de la pestaña Inicio de la cinta enumera los estilos de párrafo
del DOCX cargado. El editor resuelve la alineación, la dirección, el espaciado y la
sangría compatibles a través de los valores predeterminados del documento y de la
cadena de herencia del estilo seleccionado. El formato directo de párrafo sigue teniendo
prioridad. Elegir un estilo cambia su referencia, sin aplanar las propiedades heredadas
en cada párrafo.

La parte de estilos original se conserva al guardar el documento cargado. Todavía no se
admite crear ni editar definiciones de estilo, estilos de carácter, formato de fragmento
heredado ni la resolución de temas. Se rechaza la exportación de un paquete nuevo cuyo
catálogo de estilos difiere de los valores predeterminados integrados; use Guardar sobre el documento cargado para conservar sus
definiciones de estilo originales.
