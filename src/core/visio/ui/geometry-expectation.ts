import type { VisioDocument } from '../model';
import type { VisioResizeFrame } from './shape-resize';

/** Physical page-space poses a gesture promises, checked before accepting an edited source. */
export interface VisioGeometryExpectation {
	pageId: string;
	members: readonly { id: string; frame: VisioResizeFrame }[];
}

/** Own the small primitive snapshot across an asynchronous edit, without retaining scene trees. */
export function snapshotVisioGeometryExpectation(
	expected: VisioGeometryExpectation,
): VisioGeometryExpectation {
	if (
		!expected.members.length ||
		expected.members.length > 500 ||
		new Set(expected.members.map((member) => member.id)).size !== expected.members.length
	)
		throw new Error('Invalid selection geometry expectation.');
	return {
		pageId: expected.pageId,
		members: expected.members.map(({ id, frame }) => {
			if (
				frame.transform.length !== 6 ||
				![frame.width, frame.height, frame.pinX, frame.pinY, ...frame.transform].every(
					Number.isFinite,
				) ||
				frame.width <= 0 ||
				frame.height <= 0
			)
				throw new Error('Invalid selection geometry expectation.');
			return { id, frame: { ...frame, transform: [...frame.transform] } };
		}),
	};
}

/** A later member's formula recalculation can change an earlier member after its own command
 * succeeded. Reject that whole result instead of committing geometry different from the preview.
 */
export function assertVisioGeometryExpectation(
	document: VisioDocument,
	expected: VisioGeometryExpectation,
): void {
	const pages = document.pages.filter((page) => page.id === expected.pageId);
	const page = pages.length === 1 ? pages[0] : undefined;
	const close = (actual: number, wanted: number) =>
		Number.isFinite(actual) && Math.abs(actual - wanted) <= 1e-8 * Math.max(1, Math.abs(wanted));
	for (const { id, frame } of expected.members) {
		const shapes = page?.shapes.filter((shape) => shape.id === id);
		const shape = shapes?.length === 1 ? shapes[0] : undefined;
		if (
			!shape?.rotation ||
			!close(shape.width, frame.width) ||
			!close(shape.height, frame.height) ||
			!close(shape.rotation.pinX, frame.pinX) ||
			!close(shape.rotation.pinY, frame.pinY) ||
			shape.transform.some((value, index) => !close(value, frame.transform[index]!))
		)
			throw new Error(
				'Dependent formulas changed the requested selection geometry. The resize was not applied.',
			);
	}
}
