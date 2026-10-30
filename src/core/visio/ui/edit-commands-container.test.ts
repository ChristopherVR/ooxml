import { describe, expect, it } from 'vitest';
import type { VisioEdit } from '../index';
import { snapshotEdits } from './edit-commands';

describe('snapshotEdits for Container and Picture Format edits', () => {
	it('accepts the container options, fit, picture adjust and crop edits the ribbon sends', () => {
		const target = { pageId: '0', shapeId: '7' };
		const edits: VisioEdit[] = [
			{ type: 'set-container-options', ...target, locked: true },
			{ type: 'set-container-options', ...target, margin: 0.25, resize: 2 },
			{ type: 'fit-container', ...target },
			{ type: 'set-picture-adjust', ...target, brightness: 0.7 },
			{ type: 'crop-picture', ...target, crop: 'square' },
		];
		expect(snapshotEdits(edits)).toEqual(edits);
	});
});
