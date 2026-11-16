import type { DrawingColor } from './types';
export interface DrawingScene3D {
	/** Camera preset type, e.g. "orthographicFront", "perspectiveFront". */
	cameraPreset?: string | undefined;
	/** Camera field of view in 1/60000 degrees (`a:camera/@fov`). */
	cameraFieldOfView?: number | undefined;
	/** Camera zoom as an OOXML percentage fraction (`a:camera/@zoom`, 1 = 100%). */
	cameraZoom?: number | undefined;
	/** Camera rotation around X axis in 1/60000 degrees. */
	cameraRotX?: number | undefined;
	/** Camera rotation around Y axis in 1/60000 degrees. */
	cameraRotY?: number | undefined;
	/** Camera rotation around Z axis in 1/60000 degrees. */
	cameraRotZ?: number | undefined;
	/** Light rig type, e.g. "threePt", "balanced", "harsh". */
	lightRigType?: string | undefined;
	/** Light rig direction, e.g. "t", "b", "l", "r", "tl". */
	lightRigDirection?: string | undefined;
	/** Light-rig rotation latitude in 1/60000 degrees. */
	lightRigRotX?: number | undefined;
	/** Light-rig rotation longitude in 1/60000 degrees. */
	lightRigRotY?: number | undefined;
	/** Light-rig rotation revolution in 1/60000 degrees. */
	lightRigRotZ?: number | undefined;
	/** Whether a 3D backdrop plane is present (`a:backdrop`). */
	hasBackdrop?: boolean | undefined;
	/** Backdrop plane anchor X in EMU. */
	backdropAnchorX?: number | undefined;
	/** Backdrop plane anchor Y in EMU. */
	backdropAnchorY?: number | undefined;
	/** Backdrop plane anchor Z in EMU. */
	backdropAnchorZ?: number | undefined;
	/** Backdrop normal vector X component. */
	backdropNormalX?: number | undefined;
	/** Backdrop normal vector Y component. */
	backdropNormalY?: number | undefined;
	/** Backdrop normal vector Z component. */
	backdropNormalZ?: number | undefined;
	/** Backdrop up vector X component. */
	backdropUpX?: number | undefined;
	/** Backdrop up vector Y component. */
	backdropUpY?: number | undefined;
	/** Backdrop up vector Z component. */
	backdropUpZ?: number | undefined;
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
