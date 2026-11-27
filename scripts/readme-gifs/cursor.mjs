/**
 * A visible pointer for the README recordings. Headless Chromium draws no cursor, so this init
 * script (run inside the page, self-contained) follows the mouse with an arrow and rings each
 * click. It lives in the top layer as a manual popover, shown again after every click so it stays
 * above dialogs and menus opened since.
 */
export function cursorInit() {
	const ARROW =
		'<svg width="22" height="26" viewBox="0 0 22 26" xmlns="http://www.w3.org/2000/svg">' +
		'<path d="M2 2 L2 21 L7 16.5 L10.5 24 L14 22.5 L10.6 15.2 L17.5 15.2 Z" fill="#fff" ' +
		'stroke="#1b1b1b" stroke-width="1.6" stroke-linejoin="round"/></svg>';
	const install = () => {
		const cursor = document.createElement('div');
		cursor.setAttribute('popover', 'manual');
		cursor.setAttribute('aria-hidden', 'true');
		cursor.innerHTML = ARROW + '<span></span>';
		Object.assign(cursor.style, {
			position: 'fixed',
			inset: 'auto',
			left: '0',
			top: '0',
			margin: '0',
			padding: '0',
			border: '0',
			background: 'transparent',
			overflow: 'visible',
			pointerEvents: 'none',
			width: '22px',
			height: '26px',
			transform: 'translate(-100px, -100px)',
			filter: 'drop-shadow(0 1px 2px rgba(0,0,0,.35))',
		});
		const ring = cursor.querySelector('span');
		Object.assign(ring.style, {
			position: 'absolute',
			left: '-14px',
			top: '-14px',
			width: '28px',
			height: '28px',
			borderRadius: '50%',
			background: 'rgba(15,108,189,.35)',
			border: '2px solid rgba(15,108,189,.9)',
			opacity: '0',
			transform: 'scale(.4)',
			transition: 'opacity .25s, transform .25s',
		});
		document.documentElement.append(cursor);
		const raise = () => {
			try {
				if (cursor.matches(':popover-open')) cursor.hidePopover();
				cursor.showPopover();
			} catch {
				// Popovers are unavailable: the arrow stays a fixed element.
			}
		};
		raise();
		addEventListener(
			'mousemove',
			(event) => {
				cursor.style.transform = `translate(${event.clientX}px, ${event.clientY}px)`;
			},
			true,
		);
		addEventListener(
			'mousedown',
			() => {
				ring.style.opacity = '1';
				ring.style.transform = 'scale(1)';
			},
			true,
		);
		addEventListener(
			'mouseup',
			() => {
				ring.style.opacity = '0';
				ring.style.transform = 'scale(.4)';
				setTimeout(raise, 30);
			},
			true,
		);
		// Menus and dialogs opened by the keyboard or a timer join the top layer above the arrow.
		setInterval(raise, 400);
	};
	if (document.documentElement) install();
	else addEventListener('DOMContentLoaded', install, { once: true });
}
