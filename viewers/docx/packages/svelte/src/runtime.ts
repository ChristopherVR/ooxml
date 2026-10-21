// Plain-JavaScript half of the Svelte package: `WordEditor.svelte` imports its helpers from here.
export * from 'docx-bindings';
export * from 'docx-bindings/common';

export * from './building-blocks';

// Explicit export resolves the shared star exports in production bundlers.
export { mountEditorSurface } from '../../bindings/src/index';
