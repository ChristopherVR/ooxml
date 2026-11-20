import { describe, expect, it } from 'vitest';

import { createRoomJoinTracker } from './collaboration-room-join';

describe('room join tracker', () => {
	it('reports the first adoption after attaching as the join, and only that one', () => {
		const join = createRoomJoinTracker();
		join.attached();
		expect(join.adopted()).toBe(true);
		expect(join.adopted()).toBe(false);
	});

	it('treats room changes after this editor seeded the room as edits', () => {
		const join = createRoomJoinTracker();
		join.attached();
		join.wrote();
		expect(join.adopted()).toBe(false);
	});

	it('keeps a joining session from writing until it adopted the room slides', () => {
		const join = createRoomJoinTracker();
		join.attached('join');
		expect(join.mayWrite()).toBe(false);
		join.loaded('bootstrap');
		expect(join.mayWrite()).toBe(false);
		expect(join.adopted()).toBe(true);
		expect(join.mayWrite()).toBe(true);
	});

	it('lets a joining session publish a deck the user opened', () => {
		const join = createRoomJoinTracker();
		join.attached('join');
		join.loaded('user');
		expect(join.mayWrite()).toBe(true);
	});

	it('lets a session that started the room write at once', () => {
		const join = createRoomJoinTracker();
		join.attached();
		expect(join.mayWrite()).toBe(true);
		join.attached('create');
		expect(join.mayWrite()).toBe(true);
	});

	it('reports nothing before a session attaches, and again after the next attach', () => {
		const join = createRoomJoinTracker();
		expect(join.adopted()).toBe(false);
		join.attached();
		join.wrote();
		join.attached();
		expect(join.adopted()).toBe(true);
	});
});
