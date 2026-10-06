/**
 * The connector poly-lines of a SmartArt 3D model (cycle arcs, timeline axis and arrowheads,
 * bending-process arrows, hierarchy elbows), for `view-scene.ts`.
 *
 * `SmartArt3DModel.connectors` was built from the layout's connector paths but never drawn, so
 * every family that joins its nodes with lines rendered as disconnected shapes in 3D.
 *
 * Must not import `three` at runtime (it takes the module as a parameter).
 *
 * @module smartart-3d/connector-object
 */
import type * as THREE from 'three';

import type { SmartArt3DConnector } from '../render/smartart-3d-types';
import type { ThreeModule } from '../three-view/types';
import type { Disposable } from './flat-mesh-object';

/** Lift above z=0 so a line does not z-fight a node face that sits on the base plane. */
const CONNECTOR_Z_LIFT = 0.02;

/**
 * One `Line` per connector with at least two points. They sit on the base plane under the nodes,
 * as PowerPoint draws its connectors, so a node's face covers a line that runs behind it.
 */
export function buildConnectorLines(
	three: ThreeModule,
	connectors: readonly SmartArt3DConnector[],
	disposables: Disposable[],
): THREE.Line[] {
	return connectors.flatMap((connector) => {
		if (connector.points.length < 2) {
			return [];
		}
		const geometry = new three.BufferGeometry().setFromPoints(
			connector.points.map((p) => new three.Vector3(p.x, p.y, p.z + CONNECTOR_Z_LIFT)),
		);
		// WebGL draws every line one device pixel wide: `connector.width` cannot be honoured.
		const material = new three.LineBasicMaterial({ color: connector.color });
		disposables.push(geometry, material);
		return [new three.Line(geometry, material)];
	});
}
