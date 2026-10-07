import { isLigatures } from '../index';
import { ligatureCss } from '../layout/index';
export { ligatureCss } from '../layout/index';

/** Explicitly sets every optional class so a direct None cancels inherited font features. */
export const ligatureStyle = (value: unknown): string =>
	isLigatures(value) ? `font-variant-ligatures:${ligatureCss(value)}` : '';
