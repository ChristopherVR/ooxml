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
/** The text box and connector tools: the full viewer's own, beside the shape drawing tools. */
export const DIAGRAM_OBJECT_TOOLS = ['text', 'connector'] as const;
export type DiagramObjectTool = (typeof DIAGRAM_OBJECT_TOOLS)[number];
/** Every non-pointer tool a standalone canvas offers, in toolbar order. */
export const DIAGRAM_TOOLS = [...DIAGRAM_DRAW_TOOLS, ...DIAGRAM_OBJECT_TOOLS] as const;
/** The active canvas tool: the pointer, a drawing tool, the text tool or the connector tool. */
export type DiagramTool = 'pointer' | DiagramDrawTool | DiagramObjectTool;
export function isDiagramTool(value: unknown): value is DiagramTool {
	return value === 'pointer' || (DIAGRAM_TOOLS as readonly unknown[]).includes(value);
}
