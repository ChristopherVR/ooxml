import { resolveDrawingColor, type ThemeEffectStyle } from 'ooxml-core/drawingml';

/** Paint the parsed DrawingML tiers as swatches, with bounded CSS dimensions. */
export function effectsSample(styles: ThemeEffectStyle[], accent: string): HTMLElement {
	const preview = document.createElement('span');
	preview.className = 'dve-design-effects';
	const px = (emu = 0) => Math.min(8, Math.max(0, emu / 9525));
	for (const style of styles.slice(0, 3)) {
		const shape = document.createElement('span');
		shape.style.backgroundColor = accent;
		const shadows: string[] = [];
		for (const effect of style.effects?.effects ?? []) {
			const paint =
				'color' in effect && effect.color
					? resolveDrawingColor(effect.color, { scheme: () => accent })
					: undefined;
			const color = paint
				? `color-mix(in srgb, ${paint.hex} ${paint.alpha * 100}%, transparent)`
				: '#0006';
			switch (effect.kind) {
				case 'outerShadow':
				case 'innerShadow': {
					const angle = ((effect.directionAngle ?? 5400000) / 10800000) * Math.PI;
					const distance = px(effect.distanceEmu);
					const x = Math.cos(angle) * distance;
					const y = Math.sin(angle) * distance;
					shadows.push(
						`${effect.kind === 'innerShadow' ? 'inset ' : ''}${x.toFixed(2)}px ${y.toFixed(2)}px ${px(effect.blurRadiusEmu)}px ${color}`,
					);
					break;
				}
				case 'glow':
					shadows.push(`0 0 ${px(effect.radiusEmu)}px 1px ${color}`);
					break;
				case 'softEdge':
				case 'blur':
					shape.style.filter = `blur(${px(effect.radiusEmu) / 2}px)`;
					break;
				case 'reflection':
					shape.classList.add('dve-design-reflection');
					break;
			}
		}
		shape.style.boxShadow = shadows.join(', ');
		preview.append(shape);
	}
	return preview;
}
