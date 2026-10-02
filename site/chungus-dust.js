/**
 * Thanos-style dust for the Chungus scene: each card is redrawn on a canvas
 * the way CSS draws it, sampled into particles and blown away left to right.
 */

/** Draw a card the way CSS draws it, so its dust matches what was on screen. */
function cardPixels(card, color, letter, name) {
	const { width: w, height: h } = card.getBoundingClientRect();
	const canvas = document.createElement('canvas');
	canvas.width = Math.ceil(w);
	canvas.height = Math.ceil(h);
	const ctx = /** @type {CanvasRenderingContext2D} */ (canvas.getContext('2d'));
	ctx.fillStyle = color;
	ctx.beginPath();
	ctx.roundRect(0, 0, w, h, w * 0.22);
	ctx.fill();
	ctx.fillStyle = '#fff';
	ctx.textAlign = 'center';
	ctx.textBaseline = 'middle';
	ctx.font = `800 ${w * 0.42}px Schibsted Grotesk, system-ui, sans-serif`;
	ctx.fillText(letter, w / 2, h * 0.44);
	ctx.font = `600 ${w * 0.12}px Schibsted Grotesk, system-ui, sans-serif`;
	ctx.fillText(name, w / 2, h * 0.8);
	return {
		data: ctx.getImageData(0, 0, canvas.width, canvas.height).data,
		w: canvas.width,
		h: canvas.height,
	};
}

/** Turn every old card into drifting dust, left to right, one card after another. */
export function disintegrate(root, suite) {
	const canvas = /** @type {HTMLCanvasElement} */ (root.querySelector('.chungus__dust'));
	const ctx = /** @type {CanvasRenderingContext2D} */ (canvas.getContext('2d'));
	const dpr = Math.min(window.devicePixelRatio || 1, 2);
	canvas.width = innerWidth * dpr;
	canvas.height = innerHeight * dpr;
	ctx.scale(dpr, dpr);

	const particles = [];
	const order = suite.map((_, i) => i).sort(() => Math.random() - 0.5);
	root.querySelectorAll('.chungus__card').forEach((card, i) => {
		const app = suite[i];
		const rect = card.getBoundingClientRect();
		const { data, w, h } = cardPixels(card, app.color, app.letter, app.name);
		const step = Math.max(2, Math.round(w / 42));
		const start = order.indexOf(i) * 260;
		for (let y = 0; y < h; y += step) {
			for (let x = 0; x < w; x += step) {
				const p = (y * w + x) * 4;
				if (data[p + 3] < 128) continue;
				particles.push({
					x: Math.round(rect.left) + x,
					y: Math.round(rect.top) + y,
					size: step,
					color: `rgb(${data[p]},${data[p + 1]},${data[p + 2]})`,
					delay: start + (x / w) * 650 + Math.random() * 550,
					vx: 50 + Math.random() * 170,
					vy: -30 - Math.random() * 110,
					phase: Math.random() * 6.28,
					life: 900 + Math.random() * 900,
				});
			}
		}
		card.classList.add('is-dust');
	});

	const begin = performance.now();
	return new Promise((resolve) => {
		const frame = (now) => {
			const elapsed = now - begin;
			ctx.clearRect(0, 0, innerWidth, innerHeight);
			let alive = 0;
			for (const p of particles) {
				const t = (elapsed - p.delay) / 1000;
				const fade = t <= 0 ? 1 : 1 - (t * 1000) / p.life;
				if (fade <= 0) continue;
				alive++;
				const drift = Math.max(t, 0);
				ctx.globalAlpha = fade;
				ctx.fillStyle = p.color;
				ctx.fillRect(
					p.x + p.vx * drift + Math.sin(drift * 4 + p.phase) * 10 * drift,
					p.y + p.vy * drift - 40 * drift * drift,
					p.size * (t > 0 ? 0.8 : 1),
					p.size * (t > 0 ? 0.8 : 1),
				);
			}
			ctx.globalAlpha = 1;
			if (alive && root.isConnected) requestAnimationFrame(frame);
			else resolve(undefined);
		};
		requestAnimationFrame(frame);
	});
}
