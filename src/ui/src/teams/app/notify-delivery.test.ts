import type { ActivityItem } from 'ooxml-core/teams';
import { type DeliveryEnv, deliverNotification } from './notify-delivery';

const item: ActivityItem = {
	id: 'mention:m1',
	kind: 'mention',
	channelId: 'ops',
	channelName: 'Ops',
	messageId: 'm1',
	rootId: 'm1',
	actorId: 'bob',
	actorName: 'Bob',
	ts: 1,
	preview: 'hi @Ada',
	read: false,
};

function fake(over: Partial<DeliveryEnv> = {}) {
	const calls: string[] = [];
	const clicks: (() => void)[] = [];
	const env: DeliveryEnv = {
		hidden: () => false,
		permission: () => 'granted',
		show: (title, options, onClick) => {
			calls.push(`desktop:${title}|${options.body}|${options.tag}`);
			clicks.push(onClick);
		},
		beep: () => void calls.push('beep'),
		...over,
	};
	const actions = {
		toast: () => void calls.push('toast'),
		open: () => void calls.push('open'),
	};
	return { env, actions, calls, clicks };
}

describe('deliverNotification', () => {
	it('shows a toast, and the sound, in a visible tab', () => {
		const { env, actions, calls } = fake();
		expect(deliverNotification(item, { sound: true }, env, actions)).toBe('toast');
		expect(calls).toEqual(['toast', 'beep']);
	});

	it('uses a desktop notification in a hidden tab and opens the item when it is clicked', () => {
		const { env, actions, calls, clicks } = fake({ hidden: () => true });
		expect(deliverNotification(item, { sound: false }, env, actions)).toBe('desktop');
		expect(calls).toEqual(['desktop:Bob mentioned you in # Ops|hi @Ada|mention:m1']);
		clicks[0]!();
		expect(calls.at(-1)).toBe('open');
	});

	it('delivers nothing in a hidden tab without permission', () => {
		for (const permission of ['default', 'denied', 'unsupported'] as const) {
			const { env, actions, calls } = fake({ hidden: () => true, permission: () => permission });
			expect(deliverNotification(item, { sound: true }, env, actions)).toBe('none');
			expect(calls).toEqual([]);
		}
	});
});
