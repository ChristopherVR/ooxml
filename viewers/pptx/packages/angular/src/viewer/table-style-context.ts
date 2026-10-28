/**
 * table-style-context.ts: lets a subtree hand its table renderers a table
 * style context instead of the one built from the enclosing deck loader.
 *
 * `TableRendererComponent` resolves table banding and theme colours from the
 * viewer-scoped `LoadContentService`. A surface that renders slides outside
 * that state (the slide thumbnail rail placed in a custom shell, fed by host
 * inputs) provides this token so its miniatures resolve tables from the same
 * theme and style map the host's canvas uses.
 *
 * @module angular-viewer/table-style-context
 */
import { InjectionToken } from '@angular/core';
import type { Signal } from '@angular/core';
import type { TableStyleContext } from 'ooxml-ui/pptx';

/**
 * Table style context for every table rendered below the provider. `undefined`
 * means "no context", the same fallback a table rendered outside the viewer gets.
 */
export const TABLE_STYLE_CONTEXT_OVERRIDE = new InjectionToken<
	Signal<TableStyleContext | undefined>
>('pptx.tableStyleContextOverride');
