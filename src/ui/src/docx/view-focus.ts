import { mountViewStyle, type ViewHost, type ViewPresentation } from './document-views';
import { translate } from './localization';
import css from './styles/view-focus.css?raw';

export function mountView(host: ViewHost): ViewPresentation {
	const { frame, element } = host;
	const style = mountViewStyle(host, css);
	const exit = document.createElement('button');
	exit.type = 'button';
	exit.className = 'dve-focus-exit';
	exit.addEventListener('click', host.leave);
	const label = () => (exit.textContent = translate(host.locale(), 'Exit Focus'));
	label();
	frame.append(exit);

	const root = element.shadowRoot;
	const fullscreen = () => root?.fullscreenElement === frame;
	let entered = false;
	let disposed = false;
	function changed() {
		if (fullscreen()) entered = true;
		else if (entered) host.leave();
	}
	frame.addEventListener('fullscreenchange', changed);
	void frame
		.requestFullscreen?.()
		.then(() => {
			if (disposed && fullscreen()) void frame.ownerDocument.exitFullscreen?.().catch(() => {});
		})
		.catch(() => {
			// The fixed frame also works when an embedding page disallows browser fullscreen.
		});

	return {
		sync: label,
		dispose() {
			disposed = true;
			frame.removeEventListener('fullscreenchange', changed);
			if (fullscreen()) void frame.ownerDocument.exitFullscreen?.().catch(() => {});
			exit.remove();
			style.remove();
		},
	};
}
