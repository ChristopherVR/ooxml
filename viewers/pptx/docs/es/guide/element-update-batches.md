---
title: Actualizaciones de elementos por lotes
description: Actualice elementos en varias diapositivas como una única operación deshacible mediante la API pública del visualizador.
---

# Actualizaciones de elementos por lotes

Use `updateElements` cuando una acción del host necesite actualizar elementos de varias diapositivas. Por ejemplo, un botón "Ajustar títulos del informe" puede reposicionar los títulos de dos páginas sin salir de la página que el usuario está editando.

El método está disponible en el controlador público del visualizador en React, Vue, Angular, Svelte y Vanilla JS. `ElementUpdate` y `ElementUpdateOptions` se exportan desde el paquete de componentes de cada enlace.

```ts
updateElements(
	updates: readonly ElementUpdate[],
	options?: ElementUpdateOptions,
): Promise<void>;

interface ElementUpdate {
	slideId: string;
	elementId: string;
	patch: Partial<PptxElement>;
}

interface ElementUpdateOptions {
	label?: string;
}
```

Aquí `viewer` es el controlador de componente público actual, y ambas diapositivas ya contienen un elemento de título:

```ts
const [first, second] = viewer.getSlides();

await viewer.updateElements(
	[
		{
			slideId: first.id,
			elementId: first.elements[0].id,
			patch: { x: 84 },
		},
		{
			slideId: second.id,
			elementId: second.elements[0].id,
			patch: { x: 90 },
		},
	],
	{ label: 'Adjust report titles' },
);

viewer.undo(); // Restore both titles in one step.
viewer.redo(); // Reapply both updates in one step.
```

## Comportamiento de confirmación e historial {#commit-and-history-behavior}

- Cada lote efectivo crea un único paso de deshacer, independiente de las ediciones anteriores y posteriores, incluido otro lote enviado en el mismo turno de JavaScript.
- Aplicar un lote mantiene la diapositiva activa y la selección de elementos. No navega a las diapositivas de destino. Deshacer y rehacer usan el comportamiento existente de restauración de instantáneas del visualizador.
- Espere a que termine antes de leer el resultado o lanzar una operación que dependa de él. En React, vuelva a obtener `ref.current` al leer a través del controlador público.
- El texto en línea pendiente se confirma por separado antes de un lote efectivo. Envíe los lotes entre interacciones de puntero completadas, no durante un arrastre o un cambio de tamaño; los cinco enlaces rechazan un lote mientras hay una interacción de puntero activa. Un lote rechazado deja intactos el texto pendiente, el estado del documento y el historial de deshacer.
- Los lotes vacíos y los lotes sin cambio neto dejan sin modificar el historial y el estado de modificado. Los lotes rechazados conservan el documento y las pilas de deshacer y rehacer.
- El método existente `updateElement` mantiene su semántica actual de edición individual en la diapositiva actual.

## Destinos, parches y errores {#targets-patches-and-errors}

Los lotes requieren un documento cargado y editable en el modo normal de edición de diapositivas. Los contextos de solo lectura, vista previa, presentación, edición de patrones y de plantillas rechazan la promesa devuelta. Los ID de destino deben resolverse de forma única: `slideId` identifica una diapositiva, y `elementId` identifica un elemento de nivel superior de esa diapositiva. Los hijos de grupos anidados y los elementos de patrón o diseño quedan fuera del alcance de este método.

Cada destino y cada parche se comprueban antes de confirmar cualquier actualización. Si un destino falta o es ambiguo, el lote entero se rechaza. Los parches deben ser objetos y no pueden reemplazar el `id` ni el `type` de un elemento. Los campos de posición, tamaño y rotación deben ser números finitos; el ancho y el alto no pueden ser negativos. Los demás campos conservan el contrato existente `Partial<PptxElement>`; esto no es un validador completo del esquema de elementos.

Los parches usan reemplazo superficial, como `updateElement`: si reemplaza una propiedad anidada, proporcione el valor anidado completo. Varias entradas que apuntan al mismo elemento se aplican en el orden de entrada. Las entradas se copian de forma defensiva, de modo que cambiar un objeto de parche después de llamar al método no altera el lote enviado.
