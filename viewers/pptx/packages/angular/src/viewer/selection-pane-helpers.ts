/**
 * selection-pane-helpers.ts: the type icons of {@link SelectionPaneComponent}. The row labels and
 * the rename decision are the shared selection pane controller's (`ooxml-ui/pptx`).
 */

/** Unicode icon by element type (no Lucide dependency in Angular). */
const ELEMENT_TYPE_ICONS: Record<string, string> = {
	text: 'T',
	shape: '▭',
	image: 'Img',
	table: '⊞',
	chart: 'Cht',
	connector: '╱',
	group: '▣',
	smartArt: '◈',
	media: '▶',
	ink: '✏',
	ole: 'OLE',
};

/** Icon glyph for an element type, with a '?' fallback. */
export function elementIcon(type: string): string {
	return ELEMENT_TYPE_ICONS[type] ?? '?';
}
