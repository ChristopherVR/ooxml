import { isLigatures } from '../index.js';
import { ligatureCss } from '../layout/index.js';
export { ligatureCss } from '../layout/index.js';

/** Explicitly sets every optional class so a direct None cancels inherited font features. */
export const ligatureStyle = (value: unknown): string =>
	isLigatures(value) ? `font-variant-ligatures:${ligatureCss(value)}` : '';
