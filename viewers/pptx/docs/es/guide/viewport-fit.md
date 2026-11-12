---
title: Ajuste al área de visualización
description: Configure el ajuste de las diapositivas dentro de un host personalizado sin cambiar las dimensiones originales de las diapositivas ni el zoom del usuario.
---

# Ajuste al área de visualización

Los cinco enlaces de interfaz aceptan las opciones de host opcionales `fitPadding` y `maxFitScale`. Use `fitPadding: 0` y `maxFitScale: null` cuando el host ya proporcione sus propios márgenes y la diapositiva deba poder crecer más allá de su tamaño original en píxeles. Si se omiten estas opciones, se conserva la política de ajuste existente.

## Opciones y valores por defecto {#options-and-defaults}

`fitPadding` es una reserva en píxeles CSS sin escalar **en cada lado**: `8` reserva 16 píxeles en cada eje; `{ horizontal: 4, vertical: 16 }` reserva 8 píxeles en horizontal y 32 en vertical. `maxFitScale` es un techo positivo para el factor de ajuste; `null` significa sin límite. Los valores no válidos recurren a los valores compartidos por defecto.

Los cinco enlaces usan los mismos valores por defecto, que son los del editor de React: un área de visualización de escritorio (768 px o más) reserva 4 / 16 px por lado y limita el factor de ajuste a 1. Un área de visualización más estrecha que 768 px reserva 12 / 12 px por lado.

| Enlace  | Relleno horizontal / vertical por lado      | Factor máximo de ajuste |
| ------- | ------------------------------------------- | ----------------------- |
| React   | 4 / 16 px (12 / 12 px por debajo de 768 px) | 1                       |
| Vue     | 4 / 16 px (12 / 12 px por debajo de 768 px) | 1                       |
| Angular | 4 / 16 px (12 / 12 px por debajo de 768 px) | 1                       |
| Svelte  | 4 / 16 px (12 / 12 px por debajo de 768 px) | 1                       |
| Vanilla | 4 / 16 px (12 / 12 px por debajo de 768 px) | 1                       |

Cada paquete de interfaz exporta los tipos `ViewportFitOptions` y `ViewportFitPadding`. Estas opciones afectan al ajuste ordinario del visualizador, no a las dimensiones originales de las diapositivas, al contenido guardado, a las miniaturas, a la exportación, al modo de presentación ni al ajuste de zoom del usuario, que es independiente. Vanilla escribe siempre el relleno resuelto en línea en su área de visualización, así que una regla de relleno de la hoja de estilos sobre ese elemento no tiene efecto; use `fitPadding` en su lugar.

## Medición {#measurement}

Todos los enlaces ajustan la diapositiva al tamaño fraccionario del área de visualización (su área de relleno sin barras de desplazamiento, con la precisión con la que el navegador la diseña), y no a los valores enteros `clientWidth` y `clientHeight`. Un área diseñada a 801.5 px se ajusta a 801.5 px, así que con `fitPadding: 0` la diapositiva la llena exactamente y nunca la sobrepasa. Los bordes se descuentan por su ancho usado, lo que mantiene su fracción con el zoom del navegador (un borde de 1 px mide 1.111 px al 90 %).

El tamaño exacto se usa solo en un contexto simple, donde nada transforma ni aplica zoom CSS al área de visualización (el zoom del navegador no cuenta). Allí, el rectángulo delimitador del área coincide con su `width` y su `height` calculados, que están en sus propios píxeles CSS. Cuando una `transform` o un `zoom` de CSS en el área o en un ancestro escala el rectángulo, la diapositiva también se diseña y pinta a través de esa escala, y su propio redondeo de diseño puede desplazar sus bordes una fracción de píxel. En ese caso el ajuste reserva un píxel, de modo que la diapositiva queda dentro del área y el área no genera desplazamiento. También se reserva un píxel mientras el área muestra una barra de desplazamiento, porque el tamaño exacto no se puede recuperar en ese caso.

Los cambios posteriores los detectan `ResizeObserver` sobre el área de contenido y sobre el área de borde del visualizador, así que un cambio solo en el relleno del área también vuelve a ajustar la diapositiva.

## Configure su visualizador {#configure-your-viewer}

Dé al host una anchura y una altura reales. Un hijo flex puede necesitar también `min-height: 0` para caber en el espacio disponible. El relleno de ajuste cero no puede recuperar el espacio ocupado por barras de herramientas, paneles laterales u otros elementos del host.

### React {#react}

```tsx
<PowerPointViewer content={content} fitPadding={0} maxFitScale={null} />
```

Se pueden pasar las mismas opciones a `useViewerBuildingBlocks` al componer `Toolbar` y `SlideCanvas` en un shell personalizado. Por ejemplo:

```tsx
const { canvasProps, toolbarProps } = useViewerBuildingBlocks({
	content,
	canEdit: true,
	fitPadding: 0,
	maxFitScale: null,
});
```

Renderice las propiedades devueltas como de costumbre; el ajuste usa el área de visualización medida del lienzo, no todo el host, incluida su barra de herramientas. Los desplazamientos existentes de las reglas siguen aplicándose. Establezca `showRulers={false}` en un `SlideCanvas` personalizado cuando deba alinearse directamente con los bordes del área de visualización.

`SlideCanvas` sigue su área de visualización automáticamente, incluidos los desmontajes y montajes condicionales mientras el hook de bloques de construcción siga montado. Si usa `useZoomViewport` de `pptx-react-viewer/internals` con su propio elemento de área de visualización, enlácelo con `ref={zoom.setCanvasViewportNode ?? zoom.canvasViewportRef}` para que el ajuste y el zoom con la rueda nativa sigan al elemento actual. El objeto `canvasViewportRef` existente sigue disponible para leer el área de visualización.

### Vue {#vue}

```vue
<PowerPointViewer :content="content" :fit-padding="0" :max-fit-scale="null" />
```

Las mismas propiedades están disponibles en `SlideCanvas`.

### Angular {#angular}

```html
<pptx-viewer [content]="content()" [fitPadding]="0" [maxFitScale]="null" />
```

### Svelte {#svelte}

```svelte
<PowerPointViewer {source} fitPadding={0} maxFitScale={null} />
```

### Vanilla JS {#vanilla-js}

```ts
const viewer = createPptxViewer(container, {
	source,
	fitPadding: 0,
	maxFitScale: null,
});
```

La disposición existente de las reglas se conserva en cada enlace. Desactive las reglas cuando el host necesite un ajuste alineado con los bordes; cambiar `fitPadding` no elimina el espacio de las reglas ni cambia los demás controles de disposición del visualizador.
