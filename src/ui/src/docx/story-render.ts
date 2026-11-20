// Read-only rendering of story blocks (headers, footers, notes, review previews and the content of
// text boxes) through the editor schema, and the drawing of shapes, which both the editing surface's
// node view and these previews use: a story preview runs no node views.
import { DOMSerializer, type Node as ProseMirrorNode } from 'prosemirror-model';
import { shapePaint } from 'ooxml-core/docx/layout';
import {
	EMU_PER_PIXEL,
	type Block,
	type DocumentModel,
	type PicturePlacement,
	type ReviewDisplayMode,
	type ShapeSpec,
	type TextBoxContent,
	type ThemeCatalog,
} from 'ooxml-core/docx';
import { paragraphAttrs, placementClass } from 'ooxml-core/docx/ui';
import { schema } from './schema';
import { runToInlineNodes } from './run-adapter';
import { stylePreviewRuns } from './preview-run-styles';
import { paintPicture, type PictureSource } from './picture-paint';

/** How story blocks are rendered. */
export interface StoryRenderOptions {
	/** Package pictures to URLs, through the editor's media cache; without it pictures show empty. */
	pictureUrl?: PictureSource;
	/**
	 * Draft view's header, footer and note previews: floating (`wp:anchor`) drawings are left out, as
	 * Word's Draft view leaves them out; inline drawings are shown.
	 */
	omitFloating?: boolean;
}

function parse<T>(value: unknown): T | undefined {
	if (typeof value !== 'string') return undefined;
	try {
		return JSON.parse(value) as T;
	} catch {
		return undefined;
	}
}

function blockNode(block: Block): ReturnType<typeof schema.node> {
	if (block.type === 'paragraph') {
		const children = block.runs.flatMap((run) => runToInlineNodes(run));
		return schema.node('paragraph', paragraphAttrs(block), children);
	}
	const rows = block.rows.map((row) =>
		schema.node(
			'tableRow',
			null,
			row.map((cell) =>
				schema.node(
					'tableCell',
					null,
					cell.paragraphs.map((paragraph) => blockNode(paragraph)),
				),
			),
		),
	);
	return schema.node('table', { id: block.id, structureEditable: false }, rows);
}

/**
 * The schema's serializer, with drawing shapes and pictures drawn as they are on the editing
 * surface (pictures from `options.pictureUrl`).
 */
function storySerializer(
	model: DocumentModel | undefined,
	options: StoryRenderOptions,
): DOMSerializer {
	const base = DOMSerializer.fromSchema(schema);
	const image = base.nodes.image!;
	const drawing = (node: ProseMirrorNode): HTMLElement => {
		if (options.omitFloating && node.attrs.anchored) {
			const omitted = document.createElement('span');
			omitted.className = 'dve-floating-omitted';
			omitted.hidden = true;
			return omitted;
		}
		if (node.attrs.shape) return shapeElement(node, model, options.pictureUrl);
		const content = DOMSerializer.renderSpec(document, image(node)).dom as HTMLElement;
		if (content.tagName !== 'IMG' || !options.pictureUrl) return content;
		const frame = document.createElement('span');
		frame.className = ['dve-picture', placementClass(node.attrs.placement)]
			.filter(Boolean)
			.join(' ');
		frame.append(content);
		paintPicture(content, node, options.pictureUrl);
		return frame;
	};
	return new DOMSerializer({ ...base.nodes, image: drawing }, base.marks);
}

/**
 * A Draft view preview of a header, footer or note: its blocks read-only, pictures through the
 * editor's media cache, and floating drawings left out as Word's Draft view leaves them out.
 */
export function renderStoryPreview(
	blocks: Block[],
	model: DocumentModel | undefined,
	mode: ReviewDisplayMode = 'all',
	pictureUrl?: PictureSource,
): DocumentFragment {
	return renderBlocks(blocks, model, mode, {
		omitFloating: true,
		...(pictureUrl ? { pictureUrl } : {}),
	});
}

/** Story blocks as read-only DOM, with styles resolved through `model` when it is given. */
export function renderBlocks(
	blocks: Block[],
	model?: DocumentModel,
	mode: ReviewDisplayMode = 'all',
	options: StoryRenderOptions = {},
): DocumentFragment {
	const serializer = storySerializer(model, options);
	const fragment = document.createDocumentFragment();
	for (const block of blocks) {
		const element = serializer.serializeNode(blockNode(block)) as HTMLElement;
		element.dataset.reviewDisplay = mode;
		if (model) stylePreviewRuns(element, block, model, serializer, mode);
		fragment.append(element);
	}
	return fragment;
}

/** `wps:bodyPr` insets default to 0.1 inch left and right, 0.05 inch top and bottom. */
const DEFAULT_INSETS_EMU = {
	insetLeftEmu: 91440,
	insetTopEmu: 45720,
	insetRightEmu: 91440,
	insetBottomEmu: 45720,
};

/**
 * A text box's paragraphs laid over its shape: inset as its `wps:bodyPr` says, anchored to the top,
 * middle or bottom, and clipped to the shape's extent. `color` is the text colour where runs set
 * none (`shapePaint`'s `textColor`: white on the default and dark fills).
 */
export function textBoxContentElement(
	content: TextBoxContent,
	model?: DocumentModel,
	color?: string,
	pictureUrl?: PictureSource,
): HTMLElement {
	const element = document.createElement('div');
	element.className = 'dve-text-box-content';
	if (color && /^#[0-9a-f]{6}$/i.test(color)) element.style.color = color;
	const inset = (side: keyof typeof DEFAULT_INSETS_EMU) =>
		`${Math.round(((content[side] ?? DEFAULT_INSETS_EMU[side]) / EMU_PER_PIXEL) * 100) / 100}px`;
	element.style.padding = [
		inset('insetTopEmu'),
		inset('insetRightEmu'),
		inset('insetBottomEmu'),
		inset('insetLeftEmu'),
	].join(' ');
	element.style.justifyContent =
		content.anchor === 'ctr' ? 'center' : content.anchor === 'b' ? 'flex-end' : 'flex-start';
	element.append(renderBlocks(content.blocks, model, 'all', pictureUrl ? { pictureUrl } : {}));
	// The box's paragraphs are not the story's: nothing may resolve them as editable blocks.
	for (const block of element.querySelectorAll('[data-id]')) block.removeAttribute('data-id');
	return element;
}

/** Whether a floating shape in front of or behind text can sit at its offsets inside its paragraph. */
function absolute(placement: PicturePlacement | undefined): placement is PicturePlacement {
	return (
		placement?.wrap === 'none' &&
		(placement.relativeFromV === undefined ||
			placement.relativeFromV === 'paragraph' ||
			placement.relativeFromV === 'line') &&
		placement.align === undefined &&
		placement.alignV === undefined
	);
}

/**
 * Draws a shape node into its frame (`dom`) and drawing holder: wrapping classes, offsets, size, the
 * geometry as SVG and, for a text box, its paragraphs over it.
 */
export function paintShape(
	dom: HTMLElement,
	holder: HTMLElement,
	node: ProseMirrorNode,
	theme: ThemeCatalog | undefined,
	model: DocumentModel | undefined,
	selected = false,
	pictureUrl?: PictureSource,
): void {
	const spec = parse<ShapeSpec>(node.attrs.shape);
	const placement = parse<PicturePlacement>(node.attrs.placement);
	const content = parse<TextBoxContent>(node.attrs.textBoxContent);
	const width = Number(node.attrs.widthPx);
	const height = Number(node.attrs.heightPx);
	const anchored = Boolean(node.attrs.anchored);
	const classes = ['dve-shape', selected ? 'dve-shape-selected' : ''];
	if (anchored && absolute(placement)) {
		classes.push('dve-shape-floating');
		if (placement.behindText) classes.push('dve-shape-behind');
		dom.style.left = `${placement.offsetXPx ?? 0}px`;
		dom.style.top = `${placement.offsetYPx ?? 0}px`;
	} else {
		dom.style.left = '';
		dom.style.top = '';
		if (anchored) classes.push(placementClass(JSON.stringify(placement ?? {})));
	}
	dom.className = classes.filter(Boolean).join(' ');
	dom.dataset.shapePreset = spec?.preset ?? '';
	dom.style.width = `${width}px`;
	dom.style.height = `${height}px`;
	const paint = spec
		? shapePaint(
				{
					shape: spec,
					widthPx: width,
					heightPx: height,
					...(content ? { textBoxContent: content } : {}),
				},
				theme,
			)
		: {};
	const svg = paint.chart?.svg;
	holder.replaceChildren();
	if (svg) {
		const parsed = new DOMParser().parseFromString(svg, 'image/svg+xml').documentElement;
		if (parsed.localName === 'svg') holder.append(document.importNode(parsed, true));
	} else holder.textContent = String(node.attrs.unsupported ?? 'Shape');
	if (svg && paint.textBox)
		holder.append(textBoxContentElement(paint.textBox, model, paint.chart?.textColor, pictureUrl));
}

/** A read-only shape for a story preview, drawn as the editing surface draws it. */
function shapeElement(
	node: ProseMirrorNode,
	model: DocumentModel | undefined,
	pictureUrl: PictureSource | undefined,
): HTMLElement {
	const dom = document.createElement('span');
	const holder = document.createElement('span');
	holder.className = 'dve-shape-drawing';
	dom.append(holder);
	dom.contentEditable = 'false';
	paintShape(dom, holder, node, model?.theme, model, false, pictureUrl);
	return dom;
}
