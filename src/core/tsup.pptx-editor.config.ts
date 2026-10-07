import { defineConfig } from 'tsup';

// Shared chunks preserve collaboration leases and other module state across public helpers.
export default defineConfig({
	entry: {
		'pptx/editor/ai/table-merge': 'pptx/editor/ai/table-merge.ts',
		'pptx/editor/loader/element-patch-walker': 'pptx/editor/loader/element-patch-walker.ts',
		'pptx/editor/loader/is-external-url': 'pptx/editor/loader/is-external-url.ts',
		'pptx/editor/loader/lazy-image-resolution': 'pptx/editor/loader/lazy-image-resolution.ts',
		'pptx/editor/loader/load-content-helpers': 'pptx/editor/loader/load-content-helpers.ts',
		'pptx/editor/loader/table-style-image-paths': 'pptx/editor/loader/table-style-image-paths.ts',
		'pptx/editor/loader/text-fill-image-paths': 'pptx/editor/loader/text-fill-image-paths.ts',
		'pptx/editor/render/bullet-autonum': 'pptx/editor/render/bullet-autonum.ts',
		'pptx/editor/render/bullet-list': 'pptx/editor/render/bullet-list.ts',
		'pptx/editor/render/bullet-toggle': 'pptx/editor/render/bullet-toggle.ts',
		'pptx/editor/render/chart-user-shape-edit': 'pptx/editor/render/chart-user-shape-edit.ts',
		'pptx/editor/render/clone': 'pptx/editor/render/clone.ts',
		'pptx/editor/render/collaboration-active-session':
			'pptx/editor/render/collaboration-active-session.ts',
		'pptx/editor/render/collaboration-assets': 'pptx/editor/render/collaboration-assets.ts',
		'pptx/editor/render/collaboration-departure': 'pptx/editor/render/collaboration-departure.ts',
		'pptx/editor/render/collaboration-live-patch-target':
			'pptx/editor/render/collaboration-live-patch-target.ts',
		'pptx/editor/render/collaboration-load-origin':
			'pptx/editor/render/collaboration-load-origin.ts',
		'pptx/editor/render/collaboration-reconcile': 'pptx/editor/render/collaboration-reconcile.ts',
		'pptx/editor/render/collaboration-sync': 'pptx/editor/render/collaboration-sync.ts',
		'pptx/editor/render/collaboration-text-codec': 'pptx/editor/render/collaboration-text-codec.ts',
		'pptx/editor/render/collaboration-text-lease': 'pptx/editor/render/collaboration-text-lease.ts',
		'pptx/editor/render/collaboration-text-merge': 'pptx/editor/render/collaboration-text-merge.ts',
		'pptx/editor/render/collaboration-text-native-edit':
			'pptx/editor/render/collaboration-text-native-edit.ts',
		'pptx/editor/render/collaboration-text-positions':
			'pptx/editor/render/collaboration-text-positions.ts',
		'pptx/editor/render/collaboration-text-projection':
			'pptx/editor/render/collaboration-text-projection.ts',
		'pptx/editor/render/collaboration-text-session-apply':
			'pptx/editor/render/collaboration-text-session-apply.ts',
		'pptx/editor/render/collaboration-text-session-delta':
			'pptx/editor/render/collaboration-text-session-delta.ts',
		'pptx/editor/render/collaboration-text-session-paragraph':
			'pptx/editor/render/collaboration-text-session-paragraph.ts',
		'pptx/editor/render/collaboration-text-session':
			'pptx/editor/render/collaboration-text-session.ts',
		'pptx/editor/render/collaboration-text-snapshot-positions':
			'pptx/editor/render/collaboration-text-snapshot-positions.ts',
		'pptx/editor/render/deck-save-encryption': 'pptx/editor/render/deck-save-encryption.ts',
		'pptx/editor/render/editor-history': 'pptx/editor/render/editor-history.ts',
		'pptx/editor/render/editor-mutations': 'pptx/editor/render/editor-mutations.ts',
		'pptx/editor/render/element-locks': 'pptx/editor/render/element-locks.ts',
		'pptx/editor/render/element-operations': 'pptx/editor/render/element-operations.ts',
		'pptx/editor/render/element': 'pptx/editor/render/element.ts',
		'pptx/editor/render/find-replace': 'pptx/editor/render/find-replace.ts',
		'pptx/editor/render/group-drill': 'pptx/editor/render/group-drill.ts',
		'pptx/editor/render/group-ops': 'pptx/editor/render/group-ops.ts',
		'pptx/editor/render/inline-list-body': 'pptx/editor/render/inline-list-body.ts',
		'pptx/editor/render/remap-empty-paragraph': 'pptx/editor/render/remap-empty-paragraph.ts',
		'pptx/editor/render/remap-paragraph-sources': 'pptx/editor/render/remap-paragraph-sources.ts',
		'pptx/editor/render/remap-text-bullets': 'pptx/editor/render/remap-text-bullets.ts',
		'pptx/editor/render/remap-text': 'pptx/editor/render/remap-text.ts',
		'pptx/editor/render/schema-label-keys': 'pptx/editor/render/schema-label-keys.ts',
		'pptx/editor/render/section-operations': 'pptx/editor/render/section-operations.ts',
		'pptx/editor/render/slide-background-patch': 'pptx/editor/render/slide-background-patch.ts',
		'pptx/editor/render/slide-operations': 'pptx/editor/render/slide-operations.ts',
		'pptx/editor/render/slide-search': 'pptx/editor/render/slide-search.ts',
		'pptx/editor/render/slide-size-rescale': 'pptx/editor/render/slide-size-rescale.ts',
		'pptx/editor/render/slide-transition-edits': 'pptx/editor/render/slide-transition-edits.ts',
		'pptx/editor/render/table-cell-edit': 'pptx/editor/render/table-cell-edit.ts',
		'pptx/editor/render/table-cell-merge': 'pptx/editor/render/table-cell-merge.ts',
		'pptx/editor/render/table-data-grid-ops': 'pptx/editor/render/table-data-grid-ops.ts',
		'pptx/editor/render/table-layout': 'pptx/editor/render/table-layout.ts',
		'pptx/editor/render/table-merge': 'pptx/editor/render/table-merge.ts',
		'pptx/editor/render/table-style-editor-descriptor':
			'pptx/editor/render/table-style-editor-descriptor.ts',
		'pptx/editor/render/table-style-editor-edit': 'pptx/editor/render/table-style-editor-edit.ts',
		'pptx/editor/render/table-style-editor-parts': 'pptx/editor/render/table-style-editor-parts.ts',
		'pptx/editor/render/table-style-map-edits': 'pptx/editor/render/table-style-map-edits.ts',
		'pptx/editor/render/template-editing': 'pptx/editor/render/template-editing.ts',
		'pptx/editor/render/text-segment-paragraph-break':
			'pptx/editor/render/text-segment-paragraph-break.ts',
		'pptx/editor/render/text-theme': 'pptx/editor/render/text-theme.ts',
		'pptx/editor/render/theme-editor-model': 'pptx/editor/render/theme-editor-model.ts',
		'pptx/editor/render/theme-editor-presets': 'pptx/editor/render/theme-editor-presets.ts',
	},
	outDir: 'dist',
	tsconfig: 'tsconfig.pptx.json',
	format: ['esm', 'cjs'],
	outExtension: ({ format }) => ({ js: format === 'esm' ? '.mjs' : '.cjs' }),
	splitting: true,
	clean: false,
	dts: false,
	treeshake: true,
	platform: 'neutral',
	external: [/^ooxml-core(?:\/|$)/, 'yjs'],
});
