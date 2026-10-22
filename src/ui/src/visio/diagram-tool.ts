/** The drawing tools a standalone diagram canvas offers; every one needs edit permission. */
export const DIAGRAM_DRAW_TOOLS = [
	'rectangle',
	'ellipse',
	'line',
	'freeform',
	'arc',
	'pencil',
] as const;
export type DiagramDrawTool = (typeof DIAGRAM_DRAW_TOOLS)[number];
/** The active canvas tool: the pointer, or one drawing tool that creates a shape on a drag. */
export type DiagramTool = 'pointer' | DiagramDrawTool;
export function isDiagramTool(value: unknown): value is DiagramTool {
	return value === 'pointer' || (DIAGRAM_DRAW_TOOLS as readonly unknown[]).includes(value);
}
