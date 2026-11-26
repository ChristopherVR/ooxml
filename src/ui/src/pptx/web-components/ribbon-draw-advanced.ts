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
import { createDrawRibbonControls } from './draw-ribbon-controls';
import { drawRulerFor } from './draw-ruler';

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
	const rulerModel = root && drawRulerFor(root);
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
	const controls = createDrawRibbonControls(doc, t);
	const { groups, group, button } = controls;
	const unavailable = () =>
		!state().editable
			? 'pptx.draw.readOnly'
			: port()?.missingCommit
				? 'pptx.draw.missingCommit'
				: !canvas?.stage
					? 'pptx.draw.noSlide'
					: '';
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
		rulerModel?.set(!rulerModel.visible, Number(angle.value));
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
		rulerModel?.set(true, Number(angle.value));
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
			announce('pptx.draw.replayUnavailable');
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
		controls.sync(replay ? play : undefined);
		lasso.setAttribute('aria-pressed', String(Boolean(canvas?.armed)));
		ruler.setAttribute('aria-pressed', String(Boolean(rulerModel?.visible)));
		angle.hidden = !rulerModel?.visible || !canvas?.stage;
		angle.value = String(rulerModel?.angle ?? 0);
		rulerModel?.setLabels(t('pptx.draw.ruler'), t('pptx.draw.rulerKeys'));
		angle.setAttribute('aria-label', t('pptx.draw.rulerAngle'));
		angle.setAttribute('aria-valuetext', `${angle.value}Ãƒâ€šÃ‚Â°`);
		seek.setAttribute('aria-label', t('pptx.draw.seek'));
		play.hidden = close.hidden = seek.hidden = !replay;
		launch.hidden = Boolean(replay);
	}
	host.append(...groups);
	const unsubscribeRuler = rulerModel?.subscribe(sync);
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
			unsubscribeRuler?.();
			listeners.abort();
			canvas?.dispose();
			stopReplay();
			stopMath();
			for (const element of groups) element.remove();
		},
	};
}
