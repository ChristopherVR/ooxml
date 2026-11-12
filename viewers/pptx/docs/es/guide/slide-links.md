# Vínculos a diapositivas

Los hipervínculos de texto y de forma que apuntan a una diapositiva navegan dentro del visualizador en modo de solo lectura y de vista previa, y también durante una presentación de diapositivas. Los vínculos internos `ppaction://` nunca abren una pestaña del navegador. Al editar, use Ctrl+clic (Cmd+clic en macOS) para seguir un vínculo interno a una diapositiva; los clics normales siguen seleccionando o editando el elemento.

Los visualizadores de solo lectura y de vista previa ocultan el indicador de acción en ámbar y su información emergente de creación. Los visualizadores de edición siguen mostrándolos.

## Callback del anfitrión {#host-callback}

Cada enlace ofrece un callback antes de seguir un vínculo de texto o de forma. Devuelva `false` para cancelar la navegación por defecto, incluida la navegación externa.

| Enlace                    | Callback                                     |
| ------------------------- | -------------------------------------------- |
| Opciones de Vanilla       | `onHyperlinkClick`                           |
| Propiedad de React        | `onHyperlinkClick`                           |
| Propiedad callback de Vue | `onHyperlinkClick` (o `:on-hyperlink-click`) |
| Input de Angular          | `[onHyperlinkClick]`                         |
| Propiedad de Svelte       | `onhyperlinkclick`                           |

El callback recibe `HyperlinkClickEvent` de `ooxml-ui/pptx`: la `url` externa, la `action` interna, un `targetSlideIndex` de base cero cuando se resuelve, y el `elementId` de la forma cuando está disponible. Los campos sin valor son undefined.

```ts
const viewer = createPptxViewer(container, {
	onHyperlinkClick: (link) => {
		if (link.targetSlideIndex !== undefined) {
			selectSlideInHost(link.targetSlideIndex);
			return false;
		}
	},
});
```

Sin callback, el visualizador sigue los vínculos normalmente. Los vínculos externos conservan el marco de destino de la presentación, las comprobaciones de seguridad de URL y el ajuste de confirmación del Centro de confianza.
