import type { VisioEdit } from '../index';
import { isVisioFormatEdit, snapshotFormatting } from '../edit-formatting-commands';
import { snapshotDuplicateShapes } from '../edit-duplicate-commands';
import { snapshotPasteShapes } from '../edit-paste-commands';
import { snapshotChangeShape } from '../edit-change-shape-commands';
import {
	isVisioMetadataEdit,
	snapshotMetadataEdit,
	snapshotPictureInsert,
} from '../edit-metadata-commands';
import { isVisioGroupEdit, snapshotGroupEdit } from '../edit-group-commands';
import { isVisioPageSetupEdit, snapshotPageSetupEdit } from '../edit-page-setup-commands';
import { snapshotResizeAnchor } from '../resize-anchor';
import { snapshotConnectorGlue } from '../edit-commands';
import { snapshotTextRanges } from '../edit-text-range-commands';
import { snapshotPathCreation } from '../edit-path-commands';
import { snapshotPageTheme } from '../edit-page-theme-commands';

/** Bound cloning and strip arbitrary host properties. Semantic validation belongs to core. */
export function snapshotEdits(edits: readonly VisioEdit[]): VisioEdit[] {
	if (!Array.isArray(edits) || edits.length > 1000)
		throw new Error('At most 1000 edits are accepted per operation.');
	let characters = 0;
	return edits.map((command): VisioEdit => {
		if (
			!command ||
			typeof command.pageId !== 'string' ||
			!command.pageId ||
			command.pageId.length > 256
		)
			throw new Error('Invalid edit command.');
		if (command.type === 'duplicate-shapes') return snapshotDuplicateShapes(command);
		if (command.type === 'paste-shapes') return snapshotPasteShapes(command);
		if (command.type === 'change-shape') return snapshotChangeShape(command);
		if (command.type === 'insert-picture') return snapshotPictureInsert(command);
		if (command.type === 'set-page-theme') return snapshotPageTheme(command);
		if (isVisioMetadataEdit(command)) return snapshotMetadataEdit(command);
		if (isVisioGroupEdit(command)) return snapshotGroupEdit(command);
		if (isVisioPageSetupEdit(command)) return snapshotPageSetupEdit(command);
		if (command.type === 'reorder-page') {
			if (!Number.isSafeInteger(command.index) || command.index < 0 || command.index > 1_000_000)
				throw new Error('Invalid page order index.');
			return { type: command.type, pageId: command.pageId, index: command.index };
		}
		if (command.type === 'delete-page') return { type: command.type, pageId: command.pageId };
		if (command.type === 'set-page-size') {
			if (
				![command.width, command.height].every(
					(value) =>
						typeof value === 'number' && Number.isFinite(value) && value > 0 && value <= 1e6,
				)
			)
				throw new Error('Invalid physical page dimensions.');
			return {
				type: command.type,
				pageId: command.pageId,
				width: command.width,
				height: command.height,
			};
		}
		const text = (value: unknown): string => {
			if (typeof value !== 'string') throw new Error('Invalid edit text.');
			characters += value.length;
			if (characters > 1_000_000)
				throw new Error('Edit text exceeds the one-million-character limit.');
			return value;
		};
		const numbers = (...values: number[]) => {
			if (values.some((value) => typeof value !== 'number' || !Number.isFinite(value)))
				throw new Error('Invalid geometry edit coordinates or dimensions.');
		};
		if (command.type === 'rename-page')
			return { type: command.type, pageId: command.pageId, name: text(command.name) };
		if (command.type === 'insert-page') {
			if (
				typeof command.afterPageId !== 'string' ||
				!command.afterPageId ||
				command.afterPageId.length > 256
			)
				throw new Error('Invalid insertion target.');
			return {
				type: command.type,
				pageId: command.pageId,
				afterPageId: command.afterPageId,
				name: text(command.name),
			};
		}
		if (typeof command.shapeId !== 'string' || !command.shapeId || command.shapeId.length > 256)
			throw new Error('Invalid edit shape target.');
		const target = { pageId: command.pageId, shapeId: command.shapeId };
		if (isVisioFormatEdit(command)) return snapshotFormatting(command);
		switch (command.type) {
			case 'create-path':
				return snapshotPathCreation(command, (message) => {
					throw new Error(message);
				});
			case 'create-text-box':
				numbers(command.x, command.y, command.width, command.height);
				return {
					type: command.type,
					...target,
					x: command.x,
					y: command.y,
					width: command.width,
					height: command.height,
					text: text(command.text),
				};
			case 'reorder-shape':
				if (!['front', 'back', 'forward', 'backward'].includes(command.order))
					throw new Error('Invalid shape order.');
				return { ...target, type: command.type, order: command.order };
			case 'create-line': {
				numbers(command.beginX, command.beginY, command.endX, command.endY);
				const connect = snapshotConnectorGlue(command.connect, command.shapeId);
				return {
					type: command.type,
					...target,
					beginX: command.beginX,
					beginY: command.beginY,
					endX: command.endX,
					endY: command.endY,
					...(connect ? { connect } : {}),
				};
			}
			case 'replace-plain-text':
				return { type: command.type, ...target, text: text(command.text) };
			case 'replace-text-ranges':
				return snapshotTextRanges(command, text);
			case 'delete-shape':
				return { type: command.type, ...target };
			case 'rotate-shape':
				numbers(command.angle);
				return { type: command.type, ...target, angle: command.angle };
			case 'flip-shape':
				if (command.axis !== 'horizontal' && command.axis !== 'vertical')
					throw new Error('Invalid flip axis.');
				return { type: command.type, ...target, axis: command.axis };
			case 'move-shape':
				numbers(command.x, command.y);
				return { type: command.type, ...target, x: command.x, y: command.y };
			case 'move-line-endpoint':
				numbers(command.x, command.y);
				if (command.endpoint !== 'begin' && command.endpoint !== 'end')
					throw new Error('Invalid line endpoint.');
				return {
					type: command.type,
					...target,
					endpoint: command.endpoint,
					x: command.x,
					y: command.y,
				};
			case 'resize-shape': {
				numbers(command.width, command.height);
				const anchor = snapshotResizeAnchor(command.anchor);
				return {
					type: command.type,
					...target,
					width: command.width,
					height: command.height,
					...(anchor === undefined ? {} : { anchor }),
				};
			}
			case 'create-rectangle':
			case 'create-ellipse':
				numbers(command.x, command.y, command.width, command.height);
				return {
					type: command.type,
					...target,
					x: command.x,
					y: command.y,
					width: command.width,
					height: command.height,
					...(command.text === undefined ? {} : { text: text(command.text) }),
					...(command.type !== 'create-rectangle' || command.shape === undefined
						? {}
						: { shape: command.shape }),
				};
			default:
				throw new Error('Invalid edit command type.');
		}
	});
}
