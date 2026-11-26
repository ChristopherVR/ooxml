import { mountViewStyle, type ViewHost, type ViewPresentation } from './document-views';
import css from './styles/view-web.css?raw';

export function mountView(host: ViewHost): ViewPresentation {
	const style = mountViewStyle(host, css);
	return { dispose: () => style.remove() };
}
