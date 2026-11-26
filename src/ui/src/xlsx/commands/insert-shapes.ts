import { SHAPE_PALETTE } from 'ooxml-core/drawingml';
import { pictureAnchorAt } from 'ooxml-core/xlsx';
import { editing, icon, target, type EditorContext } from 'ooxml-core/xlsx/ui';
import { deShapes, enShapes, esShapes, frShapes, zhShapes } from '../../ribbon/shape-names';
import { SHAPE_THUMB_VIEW_BOX, shapeThumbnailPath } from '../../ribbon/shape-thumbnail';
import { field, textArea } from '../dialogs/fields';
import { showDialog } from '../dialogs/frame';

const presets = new Map(
	SHAPE_PALETTE.flatMap((group) => group.shapes.map((shape) => [shape.preset, shape.name])),
);
const shapeNames: Record<string, Readonly<Record<string, string>>> = {
	en: enShapes,
	fr: frShapes,
	de: deShapes,
	es: esShapes,
	zh: zhShapes,
};
const translateShape = (ctx: EditorContext, name: string) =>
	shapeNames[ctx.locale().split('-')[0] ?? 'en']?.[name] ?? name;

function buildSections(names: Readonly<Record<string, string>>) {
	return SHAPE_PALETTE.map((group) => ({
		title: names[group.name] ?? group.name,
		columns: 8,
		tileWidth: 26,
		tileHeight: 26,
		items: group.shapes.map(({ preset: id, name }) => ({
			id,
			label: names[name] ?? name,
			preview: `<svg xmlns="http://www.w3.org/2000/svg" width="22" height="22" viewBox="${SHAPE_THUMB_VIEW_BOX}"><path d="${shapeThumbnailPath(id)}" fill="none" stroke="currentColor"/></svg>`,
		})),
	}));
}

const sectionsByLanguage = new Map<
	Readonly<Record<string, string>>,
	ReturnType<typeof buildSections>
>();

export function shapeSections(ctx: EditorContext) {
	const names = shapeNames[ctx.locale().split('-')[0] ?? 'en'] ?? enShapes;
	const saved = sectionsByLanguage.get(names);
	if (saved) return saved;
	const sections = buildSections(names);
	sectionsByLanguage.set(names, sections);
	return sections;
}

function textBox(ctx: EditorContext, edit: boolean) {
	const t = target(ctx);
	if (!t) return;
	const index = edit ? ctx.selection.get().drawing : undefined;
	const drawing = index === undefined ? undefined : t.ws.drawings[index];
	const shape = drawing?.kind === 'unsupported' ? drawing.shape : undefined;
	if (edit && (!shape?.textBox || shape.wordArt || shape.equationParagraphs)) return;
	const input = textArea(ctx, shape?.text?.text ?? '', 4);
	return showDialog(ctx, {
		name: 'text-box',
		heading: edit ? 'Edit Text Box' : 'Text Box',
		body: field(ctx, 'Text', input),
		opened: () => input.select(),
		submit: () => {
			const live = target(ctx);
			const command = edit ? 'drawing.edit-text-box' : 'insert.text-box';
			if (live?.session !== t.session || live.ws !== t.ws || !ctx.commands.isEnabled(command))
				return false;
			if (index !== undefined) {
				if (t.ws.drawings[index] !== drawing) return false;
				t.session.setTextBoxText(live.sheet, index, input.value);
			} else {
				const added = t.session.addShape(live.sheet, pictureAnchorAt(t.active, 180, 80), {
					kind: 'textBox',
					name: ctx.t('Text Box'),
					text: input.value,
				});
				ctx.selection.set({ drawing: added });
			}
			return true;
		},
	});
}

export function insertShapeCommands() {
	return [
		editing({
			id: 'insert.shapes',
			label: 'Shapes',
			icon: icon('shapes'),
			lock: 'objects',
			run: (ctx, arg) => {
				const t = target(ctx);
				const name = typeof arg === 'string' ? presets.get(arg) : undefined;
				if (!t || !name || typeof arg !== 'string') return;
				const size = 100;
				const index = t.session.addShape(t.sheet, pictureAnchorAt(t.active, size, size * 0.7), {
					kind: 'preset',
					preset: arg,
					name: translateShape(ctx, name),
				});
				ctx.selection.set({ drawing: index });
			},
		}),
		...([false, true] as const).map((edit) =>
			editing({
				id: edit ? 'drawing.edit-equation' : 'insert.equation',
				label: edit ? 'Edit Equation' : 'Equation',
				icon: icon('equation'),
				lock: 'objects',
				unavailable: edit
					? 'Select an editable equation text box on an editable sheet.'
					: 'An editable sheet that allows objects is required.',
				enabled: (ctx) => {
					if (!edit) return true;
					const index = ctx.selection.get().drawing;
					const drawing = index === undefined ? undefined : target(ctx)?.ws.drawings[index];
					return (
						drawing?.kind === 'unsupported' &&
						Boolean(
							drawing.shape?.textBox &&
							drawing.shape.equationParagraphs?.some((p) =>
								p.some((run) => run.kind === 'equation' && run.omml),
							),
						)
					);
				},
				run: async (ctx) => {
					const before = target(ctx);
					const selected = ctx.selection.get().drawing;
					const { openEquation } = await import('./equation');
					const live = target(ctx);
					if (
						live?.ws !== before?.ws ||
						live?.session !== before?.session ||
						(edit && selected !== ctx.selection.get().drawing)
					)
						return;
					if (ctx.commands.isEnabled(edit ? 'drawing.edit-equation' : 'insert.equation'))
						await openEquation(ctx, edit);
				},
			}),
		),
		...([false, true] as const).map((edit) =>
			editing({
				id: edit ? 'drawing.edit-text-box' : 'insert.text-box',
				label: edit ? 'Edit Text Box' : 'Text Box',
				icon: icon('textBox'),
				lock: 'objects',
				run: async (ctx) => {
					await textBox(ctx, edit);
				},
			}),
		),
	];
}
