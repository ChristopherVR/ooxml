import * as THREE from 'three';
import { RoundedBoxGeometry } from 'three/addons/geometries/RoundedBoxGeometry.js';

/**
 * A very large rabbit in a gold gauntlet, modelled from primitives for this
 * site: a pear-shaped body that is cream in front and grey behind, a small
 * head with cheek fluff, long ears and a gauntlet whose fingers can curl.
 * Units are metres-ish; he stands on y = 0 facing +z and is about 3.9 tall.
 */

/** Knuckles first (index to little finger), then the thumb and the back. */
export const STONES = [0x9b4dff, 0x2f7dff, 0xff3b4e, 0xffd23f, 0x2ed47a, 0xff8a1f];

const mat = (color, extra = {}) =>
	new THREE.MeshStandardMaterial({ color, roughness: 0.8, ...extra });

const FUR = mat(0x5d606b);
const CREAM = mat(0xf2eee6, { roughness: 0.9 });
const PINK = mat(0xf0a5b5);
const DARK = mat(0x17181d, { roughness: 0.4 });
const MOUTH = mat(0xa8182b, { roughness: 0.5 });
const LID = mat(0x6a6d78);
const NOSE = mat(0x8a4a57, { roughness: 0.45 });
const GOLD = mat(0xe3ad42, { metalness: 1, roughness: 0.28 });

function mesh(geometry, material, [x, y, z] = [0, 0, 0], [sx, sy, sz] = [1, 1, 1]) {
	const m = new THREE.Mesh(geometry, material);
	m.position.set(x, y, z);
	m.scale.set(sx, sy, sz);
	m.castShadow = true;
	return m;
}

const sphere = (r, detail = 32) => new THREE.SphereGeometry(r, detail, Math.round(detail * 0.75));
const capsule = (r, length) => new THREE.CapsuleGeometry(r, length, 8, 20);

/** The pear-shaped body, cream where it faces forward and grey round the sides. */
function body() {
	const profile = [
		[0, 0.06],
		[0.7, 0.1],
		[1.1, 0.32],
		[1.3, 0.75],
		[1.3, 1.15],
		[1.18, 1.6],
		[0.98, 2.05],
		[0.78, 2.45],
		[0.62, 2.8],
		[0.45, 3.05],
		[0, 3.15],
	].map(([r, y]) => new THREE.Vector2(r, y));
	const geometry = new THREE.LatheGeometry(new THREE.SplineCurve(profile).getPoints(40), 64);
	geometry.scale(1, 1, 0.86);
	geometry.computeVertexNormals();
	const normals = geometry.getAttribute('normal');
	const colors = [];
	const grey = new THREE.Color(0x5d606b);
	const cream = new THREE.Color(0xf2eee6);
	const c = new THREE.Color();
	for (let i = 0; i < normals.count; i++) {
		const front = THREE.MathUtils.smoothstep(normals.getZ(i), 0.62, 0.8);
		c.copy(grey).lerp(cream, front);
		colors.push(c.r, c.g, c.b);
	}
	geometry.setAttribute('color', new THREE.Float32BufferAttribute(colors, 3));
	return mesh(geometry, new THREE.MeshStandardMaterial({ vertexColors: true, roughness: 0.85 }));
}

function head() {
	const group = new THREE.Group();
	group.position.set(0, 3.12, 0.05);
	group.add(mesh(sphere(0.6), FUR, [0, 0, 0], [1, 0.9, 0.88]));
	// Cheek fluff and a muzzle that runs down into the white chest.
	group.add(mesh(sphere(0.3), CREAM, [-0.25, -0.22, 0.22], [1.15, 0.85, 0.85]));
	group.add(mesh(sphere(0.3), CREAM, [0.25, -0.22, 0.22], [1.15, 0.85, 0.85]));
	group.add(mesh(sphere(0.4), CREAM, [0, -0.3, 0.2], [1.15, 0.75, 0.82]));
	// Open grin with two buck teeth.
	group.add(mesh(sphere(0.14), MOUTH, [0, -0.33, 0.49], [1.2, 0.65, 0.36]));
	const tooth = new THREE.BoxGeometry(0.06, 0.07, 0.03);
	group.add(mesh(tooth, CREAM, [-0.033, -0.265, 0.56]));
	group.add(mesh(tooth, CREAM, [0.033, -0.265, 0.56]));
	group.add(mesh(sphere(0.055), NOSE, [0, -0.13, 0.55], [1.3, 0.85, 0.8]));
	for (const x of [-0.17, 0.17]) {
		group.add(mesh(sphere(0.15), CREAM, [x, 0.08, 0.47], [0.8, 1.0, 0.5]));
		group.add(mesh(sphere(0.046), DARK, [x * 0.92, 0.02, 0.53], [1, 1.15, 0.5]));
		// Heavy half-closed lids give him the smug look.
		const lid = mesh(
			new THREE.SphereGeometry(0.158, 24, 12, 0, Math.PI * 2, 0, Math.PI / 2),
			LID,
			[x, 0.08, 0.47],
			[0.84, 1.04, 0.66],
		);
		lid.rotation.x = -0.12;
		group.add(lid);
	}
	const ears = [-1, 1].map((side) => {
		const pivot = new THREE.Group();
		pivot.position.set(side * 0.24, 0.36, -0.08);
		pivot.rotation.z = -side * 0.1;
		pivot.add(mesh(capsule(0.15, 0.95), FUR, [0, 0.66, 0], [1, 1, 0.75]));
		pivot.add(mesh(capsule(0.09, 0.8), PINK, [0, 0.64, 0.09], [1, 1, 0.3]));
		group.add(pivot);
		return pivot;
	});
	return { group, ears };
}

/**
 * The gauntlet, fingers along -y, back of the hand towards +z and the thumb on
 * the -x side. Fingers have three joints and the thumb two, each a pivot
 * group, so the timeline can press thumb to middle finger and snap.
 */
function joint(parent, y, length, radius) {
	const pivot = new THREE.Group();
	pivot.position.y = y;
	pivot.add(mesh(capsule(radius, length), GOLD, [0, -length / 2 - radius * 0.4, 0]));
	parent.add(pivot);
	return pivot;
}

function gauntlet() {
	const hand = new THREE.Group();
	hand.add(mesh(new THREE.CylinderGeometry(0.25, 0.29, 0.36, 32), GOLD, [0, 0.12, 0]));
	for (const y of [0.25, -0.03]) {
		const rim = mesh(new THREE.TorusGeometry(0.27, 0.028, 12, 40), GOLD, [0, y, 0]);
		rim.rotation.x = Math.PI / 2;
		hand.add(rim);
	}
	hand.add(mesh(new RoundedBoxGeometry(0.6, 0.56, 0.28, 5, 0.11), GOLD, [0, -0.28, 0]));
	const stones = [];
	const addStone = (parent, color, position, r) => {
		const material = new THREE.MeshStandardMaterial({
			color: new THREE.Color(color).multiplyScalar(0.55),
			emissive: color,
			toneMapped: false,
			emissiveIntensity: 0.25,
			roughness: 0.08,
			metalness: 0,
		});
		stones.push(material);
		parent.add(mesh(sphere(r, 20), material, position, [1, 1, 0.7]));
	};
	const fingers = [-0.18, -0.06, 0.06, 0.18].map((x, i) => {
		const base = new THREE.Group();
		base.position.set(x, -0.52, 0.02);
		hand.add(base);
		const proximal = joint(base, 0, 0.13, 0.058);
		const middle = joint(proximal, -0.2, 0.09, 0.054);
		const distal = joint(middle, -0.16, 0.06, 0.05);
		addStone(hand, STONES[i], [x, -0.47, 0.14], 0.05);
		return [proximal, middle, distal];
	});
	const thumbBase = new THREE.Group();
	thumbBase.position.set(-0.27, -0.3, -0.08);
	hand.add(thumbBase);
	const thumb = [joint(thumbBase, 0, 0.17, 0.065)];
	thumb.push(joint(thumb[0], -0.27, 0.12, 0.06));
	addStone(thumb[0], STONES[4], [0, -0.12, 0.06], 0.04);
	addStone(hand, STONES[5], [0, -0.27, 0.145], 0.085);
	const glow = new THREE.PointLight(0xffc870, 0, 4, 1.5);
	glow.position.set(0, -0.35, 0.6);
	hand.add(glow);
	return { hand, fingers, thumb, stones, glow };
}

function arm(side, hand) {
	const pivot = new THREE.Group();
	pivot.position.set(side * 1.0, 2.05, 0.05);
	pivot.add(mesh(capsule(0.21, 0.75), FUR, [0, -0.5, 0]));
	const wrist = new THREE.Group();
	wrist.position.set(0, -1.02, 0);
	wrist.add(hand);
	pivot.add(wrist);
	return { pivot, wrist };
}

export function buildChungus() {
	const root = new THREE.Group();
	const figure = new THREE.Group();
	root.add(figure);
	figure.add(body());
	const { group: headGroup, ears } = head();
	figure.add(headGroup);

	const feet = [-1, 1].map((side) => {
		const foot = mesh(sphere(0.42), FUR, [side * 0.55, 0.13, 0.5], [0.95, 0.38, 1.45]);
		figure.add(foot);
		return foot;
	});

	const glove = mesh(sphere(0.25), CREAM, [0, 0, 0], [1, 1.1, 0.9]);
	const right = arm(-1, glove);
	right.pivot.rotation.set(-0.25, 0, -0.4);
	figure.add(right.pivot);

	const g = gauntlet();
	g.hand.scale.setScalar(1.55);
	const left = arm(1, g.hand);
	figure.add(left.pivot);

	return { root, figure, head: headGroup, ears, feet, arm: left, gauntlet: g };
}
