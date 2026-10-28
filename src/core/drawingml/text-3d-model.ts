import type { DrawingColor } from './types';
export interface DrawingScene3D {
	/** Camera preset type, e.g. "orthographicFront", "perspectiveFront". */
	cameraPreset?: string;
	/** Camera field of view in 1/60000 degrees (`a:camera/@fov`). */
	cameraFieldOfView?: number;
	/** Camera zoom as an OOXML percentage fraction (`a:camera/@zoom`, 1 = 100%). */
	cameraZoom?: number;
	/** Camera rotation around X axis in 1/60000 degrees. */
	cameraRotX?: number;
	/** Camera rotation around Y axis in 1/60000 degrees. */
	cameraRotY?: number;
	/** Camera rotation around Z axis in 1/60000 degrees. */
	cameraRotZ?: number;
	/** Light rig type, e.g. "threePt", "balanced", "harsh". */
	lightRigType?: string;
	/** Light rig direction, e.g. "t", "b", "l", "r", "tl". */
	lightRigDirection?: string;
	/** Light-rig rotation latitude in 1/60000 degrees. */
	lightRigRotX?: number;
	/** Light-rig rotation longitude in 1/60000 degrees. */
	lightRigRotY?: number;
	/** Light-rig rotation revolution in 1/60000 degrees. */
	lightRigRotZ?: number;
	/** Whether a 3D backdrop plane is present (`a:backdrop`). */
	hasBackdrop?: boolean;
	/** Backdrop plane anchor X in EMU. */
	backdropAnchorX?: number;
	/** Backdrop plane anchor Y in EMU. */
	backdropAnchorY?: number;
	/** Backdrop plane anchor Z in EMU. */
	backdropAnchorZ?: number;
	/** Backdrop normal vector X component. */
	backdropNormalX?: number;
	/** Backdrop normal vector Y component. */
	backdropNormalY?: number;
	/** Backdrop normal vector Z component. */
	backdropNormalZ?: number;
	/** Backdrop up vector X component. */
	backdropUpX?: number;
	/** Backdrop up vector Y component. */
	backdropUpY?: number;
	/** Backdrop up vector Z component. */
	backdropUpZ?: number;
}
export interface DrawingText3DValues {
	extrusionHeight?: number;
	z?: number;
	contourWidth?: number;
	presetMaterial?: string;
	bevelTopType?: string;
	bevelTopWidth?: number;
	bevelTopHeight?: number;
	bevelBottomType?: string;
	bevelBottomWidth?: number;
	bevelBottomHeight?: number;
}
/** Source-backed text-body scene and extrusion: all dimensions stay in authored EMU/angle units. */
export interface DrawingTextShape3D extends DrawingText3DValues {
	sourceXml?: string;
	extrusionColor?: DrawingColor;
	contourColor?: DrawingColor;
}
export interface DrawingTextScene3D extends DrawingScene3D {
	sourceXml?: string;
}
export interface DrawingFlatText3D {
	z?: number;
	sourceXml?: string;
}
