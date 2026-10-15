import { watchActivity } from './activity';

describe('watchActivity', () => {
	it('throttles input events and stops after unsubscribe', () => {
		let time = 0;
		let count = 0;
		const stop = watchActivity(
			document,
			() => count++,
			1000,
			() => time,
		);
		document.dispatchEvent(new Event('keydown'));
		document.dispatchEvent(new Event('pointermove'));
		expect(count).toBe(1);
		time = 1500;
		document.dispatchEvent(new Event('pointerdown'));
		expect(count).toBe(2);
		stop();
		time = 5000;
		document.dispatchEvent(new Event('keydown'));
		expect(count).toBe(2);
	});
});
