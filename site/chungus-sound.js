/**
 * Tiny synthesised sounds for the Chungus scene: footsteps, the gauntlet's
 * hum and the snap. Web Audio only, no files; every call is a no-op when audio
 * is unavailable. Created from the button's click, so browsers allow it.
 */
export function createSound() {
	/** @type {AudioContext | null} */
	let ac = null;
	try {
		ac = new AudioContext();
	} catch {}

	/** @param {(ac: AudioContext, now: number) => void} play */
	const safely = (play) => {
		if (!ac) return;
		try {
			play(ac, ac.currentTime);
		} catch {}
	};

	/** @param {AudioContext} ac @param {number} seconds @param {number} decay */
	const noise = (ac, seconds, decay) => {
		const buffer = ac.createBuffer(1, Math.ceil(ac.sampleRate * seconds), ac.sampleRate);
		const data = buffer.getChannelData(0);
		for (let i = 0; i < data.length; i++)
			data[i] = (Math.random() * 2 - 1) * Math.exp(-i / (ac.sampleRate * decay));
		const source = ac.createBufferSource();
		source.buffer = buffer;
		return source;
	};

	return {
		stomp() {
			safely((ac, now) => {
				const osc = ac.createOscillator();
				const gain = ac.createGain();
				osc.frequency.setValueAtTime(90, now);
				osc.frequency.exponentialRampToValueAtTime(35, now + 0.18);
				gain.gain.setValueAtTime(0.35, now);
				gain.gain.exponentialRampToValueAtTime(0.001, now + 0.22);
				osc.connect(gain).connect(ac.destination);
				osc.start(now);
				osc.stop(now + 0.25);
			});
		},
		hum() {
			safely((ac, now) => {
				const osc = ac.createOscillator();
				const gain = ac.createGain();
				osc.type = 'sawtooth';
				osc.frequency.setValueAtTime(110, now);
				osc.frequency.exponentialRampToValueAtTime(440, now + 1.2);
				gain.gain.setValueAtTime(0.0001, now);
				gain.gain.exponentialRampToValueAtTime(0.05, now + 0.9);
				gain.gain.exponentialRampToValueAtTime(0.0001, now + 1.25);
				osc.connect(gain).connect(ac.destination);
				osc.start(now);
				osc.stop(now + 1.3);
			});
		},
		snap() {
			safely((ac, now) => {
				const click = noise(ac, 0.08, 0.008);
				const band = ac.createBiquadFilter();
				band.type = 'bandpass';
				band.frequency.value = 2600;
				band.Q.value = 1.2;
				const gain = ac.createGain();
				gain.gain.value = 0.9;
				click.connect(band).connect(gain).connect(ac.destination);
				click.start(now);
				const boom = noise(ac, 1.2, 0.35);
				const low = ac.createBiquadFilter();
				low.type = 'lowpass';
				low.frequency.value = 180;
				const boomGain = ac.createGain();
				boomGain.gain.value = 0.5;
				boom.connect(low).connect(boomGain).connect(ac.destination);
				boom.start(now + 0.05);
			});
		},
		close() {
			void ac?.close().catch(() => {});
		},
	};
}
