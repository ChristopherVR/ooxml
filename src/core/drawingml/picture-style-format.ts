import type { DrawingEffect, DrawingEffectList, DrawingOuterShadow } from './effect-model';
import {
	PICTURE_REFLECTION_BLUR_EMU,
	PICTURE_REFLECTION_DISTANCE_EMU,
	type PictureStyleOuterShadow,
	type PictureStyleScene,
	type PictureStyleShape3d,
	type PictureStyleSpec,
} from './picture-styles-spec';
import type { DrawingTextScene3D, DrawingTextShape3D } from './text-3d-model';
import type { DrawingColor, DrawingFill, DrawingLine } from './types';

/** Reads the compact colour spelling used by the built-in gallery catalogues. */
export function galleryDrawingColor(spec: string): DrawingColor {
	const [base = '', ...tokens] = spec.trim().split(/\s+/u);
	if (!/^#[\da-f]{6}$/iu.test(base) && !/^[a-z][a-z\d]*$/iu.test(base))
		throw new Error('Invalid gallery colour');
	const transforms = tokens.map((token) => {
		const [name, value] = token.split(':');
		if (!name || value === undefined) throw new Error('Invalid gallery colour transform');
		return { name, value };
	});
	return base.startsWith('#')
		? { kind: 'srgb', value: base.slice(1).toUpperCase(), transforms }
		: { kind: 'scheme', value: base, transforms };
}

function outerShadow(spec: PictureStyleOuterShadow): DrawingOuterShadow {
	return {
		kind: 'outerShadow',
		blurRadiusEmu: spec.blur,
		color: galleryDrawingColor(spec.color),
		rotateWithShape: false,
		...(spec.dist === undefined ? {} : { distanceEmu: spec.dist }),
		...(spec.dir === undefined ? {} : { directionAngle: spec.dir }),
		...(spec.sx === undefined ? {} : { scaleX: spec.sx }),
		...(spec.sy === undefined ? {} : { scaleY: spec.sy }),
		...(spec.kx === undefined ? {} : { skewXAngle: spec.kx }),
		...(spec.ky === undefined ? {} : { skewYAngle: spec.ky }),
		...(spec.algn ? { alignment: spec.algn } : {}),
	};
}

export function pictureStyleScene3D(scene: PictureStyleScene): DrawingTextScene3D {
	return {
		cameraPreset: scene.camera,
		...(scene.fov !== undefined && { cameraFieldOfView: scene.fov }),
		...(scene.rot && {
			cameraRotX: scene.rot[0],
			cameraRotY: scene.rot[1],
			cameraRotZ: scene.rot[2],
		}),
		lightRigType: scene.rig,
		lightRigDirection: 't',
		...(scene.rigRev !== undefined && {
			lightRigRotX: 0,
			lightRigRotY: 0,
			lightRigRotZ: scene.rigRev,
		}),
	};
}

function shape3D(spec: PictureStyleShape3d): DrawingTextShape3D {
	return {
		bevelTopType: spec.bevelPrst ?? 'circle',
		bevelTopHeight: spec.bevelH,
		...(spec.bevelW === undefined ? {} : { bevelTopWidth: spec.bevelW }),
		...(spec.contourW === undefined ? {} : { contourWidth: spec.contourW }),
		...(spec.contourClr ? { contourColor: galleryDrawingColor(spec.contourClr) } : {}),
		...(spec.material ? { presetMaterial: spec.material } : {}),
		...(spec.extrusionH === undefined ? {} : { extrusionHeight: spec.extrusionH }),
		...(spec.extrusionClr ? { extrusionColor: galleryDrawingColor(spec.extrusionClr) } : {}),
	};
}

/** DrawingML frame fields authored by a built-in Picture Styles tile. */
export function pictureStyleFormat(spec: PictureStyleSpec) {
	const { line, fill, outer, innerBlur, reflection, softEdge, scene, sp3d } = spec;
	const outline: DrawingLine = line
		? {
				widthEmu: line.w,
				cap: line.cap,
				fill: { kind: 'solid', color: galleryDrawingColor(line.color) },
				...(line.dash ? { dash: 'solid' } : {}),
				...(line.miter ? { join: { kind: 'miter', limit: 800000 } } : {}),
				...(line.cmpd ? { compound: line.cmpd } : {}),
			}
		: { fill: { kind: 'none' } };
	const effects: DrawingEffect[] = [];
	if (innerBlur !== undefined)
		effects.push({
			kind: 'innerShadow',
			blurRadiusEmu: innerBlur,
			color: galleryDrawingColor('#000000'),
		});
	if (outer) effects.push(outerShadow(outer));
	if (reflection)
		effects.push({
			kind: 'reflection',
			blurRadiusEmu: PICTURE_REFLECTION_BLUR_EMU,
			distanceEmu: PICTURE_REFLECTION_DISTANCE_EMU,
			directionAngle: 5400000,
			startAlpha: reflection.stA,
			endPosition: reflection.endPos,
			scaleY: -100000,
			alignment: 'bl',
			rotateWithShape: false,
		});
	if (softEdge !== undefined) effects.push({ kind: 'softEdge', radiusEmu: softEdge });
	const effectList: DrawingEffectList = { effects };
	const frameFill: DrawingFill | null = fill
		? { kind: 'solid', color: galleryDrawingColor(fill) }
		: null;
	return {
		geometry: spec.geom,
		...(spec.adj ? { adjustments: spec.adj } : {}),
		fill: frameFill,
		line: outline,
		effects: effectList,
		scene3d: scene ? pictureStyleScene3D(scene) : null,
		shape3d: sp3d ? shape3D(sp3d) : null,
	};
}
