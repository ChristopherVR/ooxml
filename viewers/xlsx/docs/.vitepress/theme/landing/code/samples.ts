/**
 * Framework metadata + code samples for the landing page. None of this is
 * localized: package names, entry points and code are identical in every
 * locale. Samples mirror docs/bindings.md and the per-framework guides; keep
 * them in sync with those pages. The Excel packages are not published to npm
 * yet, so `entry` is the intended import path, not an install command.
 */

export interface FrameworkSample {
	id: string;
	label: string;
	/** Intended import path of the adapter (API preview, not yet on npm). */
	entry: string;
	file: string;
	docsHref: string;
	code: string;
}

export const FRAMEWORKS: FrameworkSample[] = [
	{
		id: 'react',
		label: 'React',
		entry: '@christophervr/xlsx-react-viewer',
		file: 'Editor.tsx',
		docsHref: '/frameworks/react',
		code: `import { useState } from 'react';
import { createWorkbook, SpreadsheetEditor } from '@christophervr/xlsx-react-viewer';

export function Editor() {
  const [workbook, setWorkbook] = useState(() => createWorkbook());
  return (
    <SpreadsheetEditor
      workbook={workbook}
      fileName="Budget.xlsx"
      onWorkbookChange={setWorkbook}
    />
  );
}`,
	},
	{
		id: 'vue',
		label: 'Vue 3',
		entry: 'xlsx-vue-viewer',
		file: 'Editor.vue',
		docsHref: '/frameworks/vue',
		code: `<script setup lang="ts">
import { shallowRef } from 'vue';
import { createWorkbook, SpreadsheetEditor } from 'xlsx-vue-viewer';

const workbook = shallowRef(createWorkbook());
</script>

<template>
  <SpreadsheetEditor :workbook="workbook" @workbook-change="workbook = $event" />
</template>`,
	},
	{
		id: 'angular',
		label: 'Angular',
		entry: 'xlsx-angular-viewer',
		file: 'editor.component.ts',
		docsHref: '/frameworks/angular',
		code: `import { Component } from '@angular/core';
import { createWorkbook, SpreadsheetEditorComponent } from 'xlsx-angular-viewer';

@Component({
  selector: 'app-editor',
  standalone: true,
  imports: [SpreadsheetEditorComponent],
  template: '<spreadsheet-editor [workbook]="workbook" (workbookChange)="workbook = $event" />',
})
export class EditorComponent {
  workbook = createWorkbook();
}`,
	},
	{
		id: 'svelte',
		label: 'Svelte 5',
		entry: 'xlsx-svelte-viewer',
		file: 'Editor.svelte',
		docsHref: '/frameworks/svelte',
		code: `<script lang="ts">
  import XlsxEditor from 'xlsx-svelte-viewer';
  import { createWorkbook } from 'xlsx-svelte-viewer/runtime';

  let workbook = $state.raw(createWorkbook());
</script>

<XlsxEditor {workbook} onworkbookchange={(next) => (workbook = next)} />`,
	},
	{
		id: 'solid',
		label: 'SolidJS',
		entry: 'xlsx-solid-viewer',
		file: 'Editor.tsx',
		docsHref: '/frameworks/solid',
		code: `import { createSignal } from 'solid-js';
import { createWorkbook, SpreadsheetEditor } from 'xlsx-solid-viewer';

export function Editor() {
  const [workbook, setWorkbook] = createSignal(createWorkbook());
  return <SpreadsheetEditor workbook={workbook()} onWorkbookChange={setWorkbook} />;
}`,
	},
	{
		id: 'vanilla',
		label: 'Vanilla JS',
		entry: 'xlsx-vanilla-viewer',
		file: 'main.ts',
		docsHref: '/frameworks/vanilla',
		code: `import { mountEditor } from 'xlsx-vanilla-viewer';

const editor = mountEditor(document.querySelector('#editor')!, {
  fileName: 'Budget.xlsx',
  onWorkbookChange: (workbook) => console.log(workbook.sheets.length),
  onWorkbookError: (error) => console.error(error),
});

await editor.load(bytes, 'Budget.xlsx'); // .xlsx, .xlsm, .xls or .csv bytes
const saved = await editor.save(); // an .xlsx Blob
editor.destroy();`,
	},
];

/** Load and save outside any UI: the framework-neutral workbook API. */
export const CORE_SAMPLE = `import { loadWorkbook, saveWorkbook, sheetToCsv } from 'xlsx-vanilla-viewer';

// Any editor package bundles this loader; it detects .xlsx, .xls or CSV from the bytes.
const workbook = await loadWorkbook(bytes, { fileName: 'Budget.xls' });
console.log(workbook.format, workbook.sheets.map((sheet) => sheet.name));

// Saves as .xlsx (a legacy .xls is converted), or one sheet as CSV.
const xlsx = await saveWorkbook(workbook, 'xlsx');
const csv = sheetToCsv(workbook, 0);`;
