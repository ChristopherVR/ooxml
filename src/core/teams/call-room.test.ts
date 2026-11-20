import { describe, expect, it } from 'vitest';
import { callRoomId, chatCallRoomId, chatRoomOfCall, chatRoomId } from './protocol';
import { createTeamsWorkspace } from './workspace';

describe('call rooms', () => {
	it('calls a private chat in the room a server checks against that chat', () => {
		const ws = createTeamsWorkspace({ workspaceId: 'acme', user: { id: 'ann', name: 'Ann' } });
		const chat = ws.chat.openDirect({ id: 'bob', name: 'Bob' });
		expect(chat).not.toBeNull();
		const call = ws.call(chat!.id);
		expect(call.roomId).toBe(chatCallRoomId('acme', chat!.id));
		expect(chatRoomOfCall(call.roomId)).toBe(chatRoomId('acme', chat!.id));
		expect(ws.call('general').roomId).toBe(callRoomId('acme', 'general'));
		ws.destroy();
	});
});
