export function chartLineDashArray(style: string | undefined, width: number): string | undefined {
	if (!style || style === 'solid') {
		return undefined;
	}
	const unit = Math.max(width, 1);
	if (style === 'dot' || style === 'sysDot') {
		return `${unit} ${unit * 2}`;
	}
	if (style === 'lgDash') {
		return `${unit * 6} ${unit * 3}`;
	}
	return `${unit * 3} ${unit * 2}`;
}
