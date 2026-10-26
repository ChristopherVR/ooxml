// SmartArt in the Excel editor: Insert > SmartArt, the SmartArt Design commands and the shared text
// pane (`SmartArtTextPaneController`, as in Word), wired to the core's edit session
// (`editSmartArt`, `changeSmartArtLayout`, `insertSmartArt`), so every change is one undoable step.
// The definitions a new or re-laid-out diagram needs come from `ooxml-core/diagram/builtin`,
// loaded on first use (its layouts are large). The commands and their labels are the shared
// SmartArt Design list of `ooxml-core/diagram`, which Word's tab is built from too.
import {
	DIAGRAM_DESIGN_COMMANDS,
	DIAGRAM_COLOR_ENTRIES,
	DIAGRAM_STYLE_ENTRIES,
	type DiagramDesignCommandKey,
	type DiagramInsertPreset,
} from 'ooxml-core/diagram';
import { pictureAnchorAt } from 'ooxml-core/xlsx';
import {
	activeSmartArt,
	editing,
	icon,
	target,
	type Command,
	type EditorContext,
} from 'ooxml-core/xlsx/ui';
import { SmartArtTextPaneController } from '../smartart';
import { chooseSmartArt } from './dialogs/insert-smartart';

/** Excel's default SmartArt size in CSS pixels (about 5 by 3 inches). */
const DEFAULT_SIZE = { width: 480, height: 280 };

const ICONS: Record<DiagramDesignCommandKey, Parameters<typeof icon>[0]> = {
	addAfter: 'smartArtAddShape',
	addBefore: 'smartArtAddShape',
	addAbove: 'smartArtAddShape',
	addBelow: 'smartArtAddShape',
	textPane: 'smartArtTextPane',
	promote: 'smartArtPromote',
	demote: 'smartArtDemote',
	moveUp: 'smartArtMoveUp',
	moveDown: 'smartArtMoveDown',
	changeLayout: 'smartArtLayout',
};

/** The command id of a SmartArt Design command (`smartart.addAfter`, ...). */
export const smartArtCommandId = (key: DiagramDesignCommandKey): string => `smartart.${key}`;

const controllers = new WeakMap<EditorContext, XlsxSmartArtController>();

class XlsxSmartArtController {
	readonly text: SmartArtTextPaneController;

	constructor(private readonly ctx: EditorContext) {
		this.text = new SmartArtTextPaneController(
			ctx.host.ownerDocument,
			{
				nodes: () => activeSmartArt(ctx)?.smartArt.nodes,
				edit: (edit) => {
					const t = target(ctx);
					const found = activeSmartArt(ctx);
					if (!t || !found || ctx.readOnly()) return undefined;
					const result = t.session.editSmartArt(t.sheet, found.index, edit);
					return result.changed ? (result.nodeId ?? '') : undefined;
				},
				setText: (id, text) => {
					const t = target(ctx);
					const found = activeSmartArt(ctx);
					if (t && found && !ctx.readOnly())
						t.session.editSmartArt(t.sheet, found.index, { kind: 'text', nodeId: id, text });
				},
				toggled: (open) => {
					ctx.requestRender();
					if (!open) ctx.grid()?.focus();
				},
			},
			'xve-smartart-text-pane',
		);
		this.relocalize();
	}

	run(key: DiagramDesignCommandKey): void {
		if (key === 'textPane') return this.text.setOpen(!this.text.open);
		if (key === 'changeLayout') return void this.relayout();
		this.text.runCommand(key);
		this.ctx.grid()?.focus();
	}

	private async definitions() {
		return import('ooxml-core/diagram/builtin');
	}

	private async layoutOf(preset: DiagramInsertPreset) {
		const builtin = await this.definitions();
		const layoutId = builtin.builtinLayoutIdForPreset(preset.layout);
		if (!layoutId) throw new Error('The SmartArt layouts could not be loaded.');
		return { builtin, layoutId, layoutXml: builtin.loadBuiltinSmartArtLayoutXml(layoutId) };
	}

	async insert(): Promise<void> {
		const ctx = this.ctx;
		const preset = await chooseSmartArt(ctx);
		const t = target(ctx);
		if (!preset || !t || ctx.readOnly()) return;
		const { builtin, layoutId, layoutXml } = await this.layoutOf(preset);
		const index = t.session.insertSmartArt(t.sheet, {
			anchor: pictureAnchorAt(t.active, DEFAULT_SIZE.width, DEFAULT_SIZE.height),
			layoutId,
			family: preset.family,
			layoutXml,
			colorsId: builtin.BUILTIN_SMARTART_COLORS_ID,
			colorsXml: builtin.BUILTIN_SMARTART_COLORS_XML,
			quickStyleId: builtin.BUILTIN_SMARTART_QUICK_STYLE_ID,
			quickStyleXml: builtin.BUILTIN_SMARTART_QUICK_STYLE_XML,
			// Excel starts every shape empty, showing its "[Text]" placeholder.
			items: preset.defaultItems.map(() => ''),
		});
		ctx.selection.set({ drawing: index });
		this.text.select(activeSmartArt(ctx)?.smartArt.nodes[0]?.id ?? null);
		this.text.setOpen(true);
	}

	private async relayout(): Promise<void> {
		const ctx = this.ctx;
		const preset = await chooseSmartArt(ctx);
		const t = target(ctx);
		const found = activeSmartArt(ctx);
		if (!preset || !t || !found || ctx.readOnly()) return;
		const { layoutId, layoutXml } = await this.layoutOf(preset);
		t.session.changeSmartArtLayout(t.sheet, found.index, {
			layoutId,
			family: preset.family,
			layoutXml,
		});
	}

	refresh(): void {
		this.text.sync();
	}

	relocalize(): void {
		this.text.setLabels((key) => this.ctx.t(key));
	}
}

/** Creates the editor's SmartArt controller; its text pane goes beside the grid. */
export function createSmartArtController(ctx: EditorContext) {
	const controller = new XlsxSmartArtController(ctx);
	controllers.set(ctx, controller);
	return {
		element: controller.text.pane as HTMLElement,
		refresh: () => controller.refresh(),
		relocalize: () => controller.relocalize(),
		/** A new workbook: the pane closes. */
		reset: () => controller.text.reset(),
	};
}

/** The SmartArt Design commands and Insert > SmartArt. */
export function smartArtCommands(): Command[] {
	const design = DIAGRAM_DESIGN_COMMANDS.map((command) =>
		editing({
			id: smartArtCommandId(command.key),
			label: command.label,
			icon: icon(ICONS[command.key]),
			lock: 'objects',
			enabled: (ctx) => !!activeSmartArt(ctx),
			...(command.key === 'textPane'
				? { checked: (ctx: EditorContext) => controllers.get(ctx)?.text.open ?? false }
				: {}),
			run: (ctx) => controllers.get(ctx)?.run(command.key),
		}),
	);
	return [
		...design,
		...(['colors', 'style'] as const).map((kind) =>
			editing({
				id: `smartart.${kind}`,
				label: kind === 'colors' ? 'Change Colors' : 'SmartArt Styles',
				icon: icon('smartArt'),
				lock: 'objects',
				enabled: (ctx) => !!activeSmartArt(ctx),
				run: (ctx, arg) => {
					const found = activeSmartArt(ctx);
					const t = target(ctx);
					if (!found || !t) return;
					if (kind === 'colors') {
						const choice = DIAGRAM_COLOR_ENTRIES.find((entry) => entry.id === arg);
						if (choice) t.session.formatSmartArt(t.sheet, found.index, { kind, value: choice.id });
					} else {
						const choice = DIAGRAM_STYLE_ENTRIES.find((entry) => entry.id === arg);
						if (choice) t.session.formatSmartArt(t.sheet, found.index, { kind, value: choice.id });
					}
				},
			}),
		),
		editing({
			id: 'insert.smartart',
			label: 'SmartArt',
			icon: icon('smartArt'),
			lock: 'objects',
			run: async (ctx) => {
				try {
					await controllers.get(ctx)?.insert();
				} catch (error) {
					ctx.toast(ctx.t(error instanceof Error ? error.message : String(error)), 'error');
				}
			},
		}),
	];
}
