import { convertSlideInk, replaceSlideInk } from 'ooxml-core/pptx/editor/ink';
import type { PptxElement, PptxSlide } from 'ooxml-core/pptx';

import { translationsEn } from '../i18n/translations-en';
import {
	DRAW_CANVAS_CHANGE,
	DRAW_CANVAS_SCOPE,
	drawCanvasPort,
	drawHostOptions,
} from '../render/draw-canvas';
import { newEquationElement } from '../render/editor-insert';
import type { RibbonDrawIntent, RibbonDrawViewState } from '../render/ribbon-draw-state';
import { createCanvasInkTools } from './draw-canvas-tools';
import { createDrawPens } from './draw-pens';
import { createDrawReplay } from './draw-replay';

export function mountRibbonDrawAdvanced(
	host: HTMLElement,
	state: () => RibbonDrawViewState,
	request: (intent: RibbonDrawIntent) => void,
) {
	const root = host.closest<HTMLElement>(DRAW_CANVAS_SCOPE);
	const doc = host.ownerDocument;
	const t = (key: string) => {
		const value = state().translate?.(key);
		return value && value !== key ? value : (translationsEn[key] ?? key);
	};
	const canvas = root
		? createCanvasInkTools(root, {
				onSelection: () => {
					stopMath();
					announce('');
					sync();
				},
				onDelete: () => {
					const slide = port()?.getSlide();
					if (slide) commit(slide);
				},
			})
		: undefined;
	const port = () => (root ? drawCanvasPort(root) : undefined);
	const settings = () => (root ? drawHostOptions(root) : undefined);
	const listeners = new AbortController();
	const viewport = root?.querySelector('[data-pptx-viewport]');
	const observer = viewport
		? new MutationObserver(() => {
				if (!canvas?.stage?.isConnected && port()?.getSlide()) sync();
			})
		: undefined;
	let replay: ReturnType<typeof createDrawReplay>;
	let mathRequest: AbortController | undefined;
	let pending = false;
	const groups: HTMLElement[] = [];
	const commands: Array<{ control: HTMLButtonElement; key: string; reason: () => string }> = [];
	const labelGroups: Array<{ group: HTMLElement; key: string }> = [];
	const group = (id: string, key: string) => {
		const element = doc.createElement('pptx-ui-ribbon-group');
		element.dataset.ribbonGroup = `draw.${id}`;
		const row = doc.createElement('div');
		row.className = 'advanced-row';
		element.append(row);
		groups.push(element);
		labelGroups.push({ group: element, key });
		return row;
	};
	const button = (
		parent: HTMLElement,
		id: string,
		key: string,
		reason: () => string,
		action: () => void,
	) => {
		const control = doc.createElement('button');
		control.type = 'button';
		control.dataset.ribbonControl = `draw.${id}`;
		control.dataset.pptxCompact = '';
		control.addEventListener('click', () => {
			if (!reason()) action();
		});
		commands.push({ control, key, reason });
		parent.append(control);
		return control;
	};
	const unavailable = () =>
		!state().editable ? 'pptx.draw.readOnly' : !canvas?.stage ? 'pptx.draw.noSlide' : '';
	const selectionReason = () =>
		unavailable() || (!canvas?.selection.length ? 'pptx.draw.noSelection' : '');
	const select = group('selection', 'pptx.ribbon.tool.select');
	const lasso = button(select, 'selection.lasso', 'pptx.draw.lasso', unavailable, () =>
		toggleLasso(!canvas?.armed),
	);
	const pens = createDrawPens({
		doc,
		state,
		host: settings,
		request,
		t,
		activate: (kind) => {
			if (kind === 'actionPen') toggleLasso(true);
			canvas?.pencil(kind === 'pencil');
		},
	});
	groups.push(pens.group);
	const convert = group('convert', 'pptx.draw.convert');
	button(convert, 'convert.inkToShape', 'pptx.draw.inkToShape', selectionReason, () => {
		const slide = port()?.getSlide();
		if (!slide || !canvas) return;
		const replacement = convertSlideInk(canvas.selection, newId());
		if (replacement) {
			commit(slide, replacement);
		} else {
			announce('pptx.draw.unrecognizedShape');
		}
	});
	const mathReason = () =>
		unavailable() ||
		(!settings()?.recognizeMath ? 'pptx.draw.mathUnavailable' : '') ||
		selectionReason() ||
		(pending ? 'pptx.draw.recognizing' : '');
	button(convert, 'convert.inkToMath', 'pptx.draw.inkToMath', mathReason, () => {
		void inkToMath();
	});
	const guides = group('guides', 'pptx.draw.guides');
	const ruler = button(guides, 'guides.ruler', 'pptx.draw.ruler', unavailable, () => {
		canvas?.ruler(!canvas.rulerVisible, Number(angle.value));
		sync();
	});
	const angle = doc.createElement('input');
	angle.type = 'range';
	angle.min = '-180';
	angle.max = '180';
	angle.step = '1';
	angle.value = '0';
	angle.dataset.ribbonControl = 'draw.guides.angle';
	angle.dataset.pptxCompact = '';
	angle.addEventListener('input', () => {
		canvas?.ruler(true, Number(angle.value));
		sync();
	});
	guides.append(angle);
	const playback = group('replay', 'pptx.draw.replay');
	const canReplay = () => unavailable() || (!canvas?.hasInk ? 'pptx.draw.noInk' : '');
	const launch = button(playback, 'replay.start', 'pptx.draw.replay', canReplay, () => {
		const stage = canvas?.stage;
		if (!stage || !canvas) return;
		stopReplay();
		canvas.arm(false);
		request({ kind: 'tool', value: 'select' });
		replay = createDrawReplay(stage, canvas.all(), (time, duration, playing) => {
			seek.max = String(duration);
			seek.value = String(time);
			play.textContent = t(playing ? 'pptx.animations.pause' : 'pptx.animations.play');
		});
		if (!replay) {
			announce('pptx.draw.noInk');
			return;
		}
		replay.toggle();
		sync();
		play.focus();
	});
	const play = button(
		playback,
		'replay.play',
		'pptx.animations.play',
		() => '',
		() => replay?.toggle(),
	);
	const seek = doc.createElement('input');
	seek.type = 'range';
	seek.min = '0';
	seek.step = '1';
	seek.dataset.ribbonControl = 'draw.replay.seek';
	seek.dataset.pptxCompact = '';
	seek.addEventListener('input', () => replay?.seek(Number(seek.value)));
	playback.append(seek);
	const close = button(
		playback,
		'replay.close',
		'pptx.draw.closeReplay',
		() => '',
		() => {
			stopReplay();
			sync();
			launch.focus();
		},
	);
	const status = doc.createElement('span');
	status.className = 'advanced-status';
	status.setAttribute('role', 'status');
	playback.append(status);
	function announce(key: string) {
		status.textContent = key ? t(key) : '';
	}
	function stopReplay() {
		replay?.dispose();
		replay = undefined;
	}
	function stopMath() {
		mathRequest?.abort();
		mathRequest = undefined;
		pending = false;
	}
	function toggleLasso(armed: boolean) {
		stopReplay();
		request({ kind: 'tool', value: 'select' });
		canvas?.arm(armed);
		sync();
	}
	function newId() {
		return `ink-conversion-${doc.defaultView?.crypto.randomUUID() ?? Date.now()}`;
	}
	function commit(slide: PptxSlide, replacement?: PptxElement) {
		const target = port();
		if (
			!canvas?.selection.length ||
			!state().editable ||
			target?.getSlide()?.elements !== slide.elements
		)
			return;
		const elements = replaceSlideInk(slide.elements, canvas.selection, replacement);
		canvas.arm(false);
		target?.replaceElements(slide, elements, replacement?.id ?? null);
		sync();
	}
	async function inkToMath() {
		const slide = port()?.getSlide();
		const provider = settings()?.recognizeMath;
		const strokes = canvas?.selection;
		const start = strokes?.[0]?.points[0];
		if (!root || !slide || !provider || !strokes || !start) return;
		stopMath();
		const abort = new AbortController();
		mathRequest = abort;
		pending = true;
		announce('pptx.draw.recognizing');
		sync();
		try {
			const { recognizeDrawEquation } = await import('./draw-math');
			const isCurrent = () =>
				!abort.signal.aborted &&
				settings()?.recognizeMath === provider &&
				port()?.getSlide()?.elements === slide.elements;
			const opened = await recognizeDrawEquation({
				root,
				t,
				provider,
				signal: abort.signal,
				isCurrent,
				strokes: strokes.map(({ points }) => points),
				insert: (omml) => {
					if (isCurrent() && canvas?.selection === strokes) {
						const equation = newEquationElement(start.x, start.y, {
							omml,
							label: t('pptx.equation.placeholderText'),
						});
						commit(slide, { ...equation, id: newId() });
					}
				},
			});
			if (isCurrent()) announce(opened ? '' : 'pptx.draw.unrecognizedMath');
		} catch {
			if (!abort.signal.aborted) announce('pptx.draw.unrecognizedMath');
		} finally {
			if (mathRequest === abort) {
				pending = false;
				sync();
			}
		}
	}
	function sync() {
		if (canvas?.sync(state().editable ? port()?.getSlide() : undefined)) {
			stopReplay();
			stopMath();
			announce('');
		}
		pens.sync();
		for (const { group: element, key } of labelGroups) element.setAttribute('label', t(key));
		for (const { control, key, reason } of commands) {
			const why = reason();
			if (control !== play || !replay) control.textContent = t(key);
			control.setAttribute('aria-disabled', String(Boolean(why)));
			control.title = t(why || key);
			control.setAttribute('aria-description', why ? t(why) : '');
		}
		lasso.setAttribute('aria-pressed', String(Boolean(canvas?.armed)));
		ruler.setAttribute('aria-pressed', String(Boolean(canvas?.rulerVisible)));
		angle.hidden = !canvas?.rulerVisible || !canvas.stage;
		angle.setAttribute('aria-label', t('pptx.draw.rulerAngle'));
		angle.setAttribute('aria-valuetext', `${angle.value}°`);
		seek.setAttribute('aria-label', t('pptx.draw.seek'));
		play.hidden = close.hidden = seek.hidden = !replay;
		launch.hidden = Boolean(replay);
	}
	host.append(...groups);
	root?.addEventListener(DRAW_CANVAS_CHANGE, sync, { signal: listeners.signal });
	if (viewport) observer?.observe(viewport, { childList: true, subtree: true });
	sync();
	return {
		sync,
		toolChanged() {
			canvas?.arm(false);
			canvas?.pencil(false);
			stopReplay();
			stopMath();
			announce('');
		},
		dispose() {
			observer?.disconnect();
			listeners.abort();
			canvas?.dispose();
			stopReplay();
			stopMath();
			for (const element of groups) element.remove();
		},
	};
}
