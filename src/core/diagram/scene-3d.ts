/**
 * The 3D a SmartArt shape and quick style carry, as values a painter can use: a shape's own
 * `a:scene3d` and `a:sp3d` (Office's cache, or the label 3D a generated cache writes), a quick-style
 * label's 3D, and the diagram's spatial camera (`dgm:styleDef/dgm:scene3d`), which Office keeps in
 * the quick-style part and not in its cached shapes. Reads through the shared DrawingML 3D readers.
 */
import { NS, first, parseXml, type XmlElement } from '../xml/index';
import {
	readScene3DValues,
	readText3DValues,
	type ThreeDReader,
} from '../drawingml/read-3d-values';
import type { DrawingScene3D, DrawingText3DValues } from '../drawingml/text-3d-model';
import type { DiagramDrawingShape, DiagramShape3dView } from './types';

const reader: ThreeDReader<XmlElement> = {
	attribute: (node, name) => node?.getAttribute(name) ?? undefined,
	child: (node, name) => (node ? first(node, name, NS.a) : undefined),
};

/** The 3D of a `dsp:spPr` (or any element holding `a:scene3d` and `a:sp3d`). */
export function readDiagramShape3d(parent: XmlElement | undefined): DiagramShape3dView | undefined {
	const sceneElement = first(parent, 'scene3d', NS.a);
	const shapeElement = first(parent, 'sp3d', NS.a);
	const view: DiagramShape3dView = {};
	if (sceneElement) view.scene = readScene3DValues(sceneElement, reader);
	if (shapeElement) {
		const shape = readText3DValues(shapeElement, reader);
		if (Object.keys(shape).length) view.shape = shape;
	}
	return view.scene || view.shape ? view : undefined;
}

/** The 3D of DrawingML `a:scene3d` / `a:sp3d` markup (a generated shape's label 3D). */
export function parseDiagramShape3dXml(xml: string): DiagramShape3dView | undefined {
	if (!xml) return undefined;
	try {
		return readDiagramShape3d(parseXml(`<x xmlns:a="${NS.a}">${xml}</x>`).documentElement);
	} catch {
		return undefined;
	}
}

/** A shape's 3D, read from Office's cache or from the label 3D of a generated shape. */
export function diagramShape3dView(shape: DiagramDrawingShape): DiagramShape3dView | undefined {
	if (shape.view3d) return shape.view3d;
	if (!shape.shape3d) return undefined;
	return parseDiagramShape3dXml(
		`${shape.shape3d.scene3dXml ?? ''}${shape.shape3d.shape3dXml ?? ''}`,
	);
}

/** Whether a scene turns the diagram away from the front view. */
export function isSpatialDiagramScene(scene: DrawingScene3D | undefined): boolean {
	if (!scene) return false;
	return Boolean(
		(scene.cameraPreset && scene.cameraPreset !== 'orthographicFront') ||
		scene.cameraRotX ||
		scene.cameraRotY ||
		scene.cameraRotZ ||
		scene.hasBackdrop,
	);
}

/** A quick-style label's 3D, by label name. */
export interface DiagramLabel3d {
	name: string;
	scene?: DrawingScene3D;
	shape?: DrawingText3DValues;
	/** The text body's own 3D (`dgm:txPr/a:sp3d`). */
	text?: DrawingText3DValues;
}

/** A quick style's 3D: its spatial camera (when it has one) and each label's 3D. */
export interface DiagramQuickStyle3d {
	scene?: DrawingScene3D;
	labels: DiagramLabel3d[];
}

/** Read a quick-style part's 3D; a flat style yields no scene and labels without 3D. */
export function diagramQuickStyle3d(quickStyleXml: string): DiagramQuickStyle3d {
	let root: XmlElement;
	try {
		root = parseXml(quickStyleXml).documentElement;
	} catch {
		return { labels: [] };
	}
	const sceneOf = (element: XmlElement | undefined) =>
		element ? readScene3DValues(element, reader) : undefined;
	const values = (element: XmlElement | undefined) => {
		if (!element) return undefined;
		const value = readText3DValues(element, reader);
		return Object.keys(value).length ? value : undefined;
	};
	const own = Array.from(root.childNodes).find(
		(node): node is XmlElement =>
			node.nodeType === 1 &&
			(node as XmlElement).localName === 'scene3d' &&
			(node as XmlElement).namespaceURI === NS.dgm,
	);
	const scene = sceneOf(own);
	const labels = Array.from(root.childNodes)
		.filter(
			(node): node is XmlElement =>
				node.nodeType === 1 &&
				(node as XmlElement).localName === 'styleLbl' &&
				(node as XmlElement).namespaceURI === NS.dgm,
		)
		.map((label) => {
			const labelScene = sceneOf(first(label, 'scene3d', NS.dgm));
			const shape = values(first(label, 'sp3d', NS.dgm));
			const text = values(first(first(label, 'txPr', NS.dgm), 'sp3d', NS.a));
			return {
				name: label.getAttribute('name') ?? '',
				...(labelScene ? { scene: labelScene } : {}),
				...(shape ? { shape } : {}),
				...(text ? { text } : {}),
			};
		});
	return { ...(isSpatialDiagramScene(scene) ? { scene } : {}), labels };
}
