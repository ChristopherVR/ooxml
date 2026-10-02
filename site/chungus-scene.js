import * as THREE from 'three';
import { RoomEnvironment } from 'three/addons/environments/RoomEnvironment.js';
import { buildChungus } from './chungus-model.js';

/**
 * The 3D stage: renderer, lights, camera and the moves the Chungus timeline
 * calls (walk in, face the camera, raise the gauntlet, charge, snap). Every
 * move returns a promise that settles when it finishes.
 */

const ease = (t) => (t < 0.5 ? 4 * t * t * t : 1 - (-2 * t + 2) ** 3 / 2);
const easeOut = (t) => 1 - (1 - t) ** 3;
const easeOutBack = (t) => 1 + 2.7 * (t - 1) ** 3 + 1.7 * (t - 1) ** 2;

const ARM_REST = new THREE.Quaternion().setFromEuler(new THREE.Euler(-1.3, 0, -0.62));
// Snapping pose: the arm reaches towards the camera at chest height and the
// wrist bends the hand upright, thumb side out, like a photo of someone snapping.
const ARM_UP = new THREE.Quaternion().setFromUnitVectors(
	new THREE.Vector3(0, -1, 0),
	new THREE.Vector3(-0.15, 0.05, 1).normalize(),
);
const WRIST_UP = { x: -1.2, y: -0.9 };

/**
 * Hand shapes: three joint curls per finger (index to little finger), then
 * the thumb's base rotation (x, y, z) and its tip curl.
 */
const HANDS = {
	fist: {
		fingers: [
			[1.5, 1.4, 0.9],
			[1.5, 1.4, 0.9],
			[1.5, 1.4, 0.9],
			[1.5, 1.4, 0.9],
		],
		thumb: [0.9, 0, 0.9, 0.6],
	},
	// Thumb pad pressed to the middle fingertip (solved so the two touch).
	ready: {
		fingers: [
			[0.3, 0.25, 0.1],
			[1.0, 1.0, 0.5],
			[1.55, 1.45, 0.9],
			[1.55, 1.45, 0.9],
		],
		thumb: [0.5, 0, 0.7, 1],
	},
	// Tension before the snap: thumb pushes, middle finger presses back.
	press: {
		fingers: [
			[0.3, 0.25, 0.1],
			[1.1, 1.05, 0.6],
			[1.55, 1.45, 0.9],
			[1.55, 1.45, 0.9],
		],
		thumb: [0.6, 0, 0.78, 1.05],
	},
	// The middle finger slips off and slams into the palm; the thumb flicks up.
	snapped: {
		fingers: [
			[0.2, 0.15, 0.05],
			[1.65, 1.55, 1.0],
			[1.55, 1.45, 0.9],
			[1.55, 1.45, 0.9],
		],
		thumb: [0.15, 0, 0.2, 0.1],
	},
};

export function createScene(host, { fast = false, onStep = () => {} } = {}) {
	const renderer = new THREE.WebGLRenderer({ alpha: true, antialias: true });
	renderer.setPixelRatio(Math.min(devicePixelRatio, 2));
	renderer.shadowMap.enabled = true;
	renderer.shadowMap.type = THREE.PCFSoftShadowMap;
	renderer.toneMapping = THREE.ACESFilmicToneMapping;
	host.append(renderer.domElement);

	const scene = new THREE.Scene();
	const pmrem = new THREE.PMREMGenerator(renderer);
	scene.environment = pmrem.fromScene(new RoomEnvironment(), 0.04).texture;
	scene.add(new THREE.HemisphereLight(0xdfe6ff, 0x3a3326, 0.7));
	const sun = new THREE.DirectionalLight(0xfff2dd, 2.2);
	sun.position.set(3, 7, 5);
	sun.castShadow = true;
	sun.shadow.mapSize.set(1024, 1024);
	Object.assign(sun.shadow.camera, { left: -5, right: 5, top: 6, bottom: -2 });
	scene.add(sun);
	const ground = new THREE.Mesh(
		new THREE.PlaneGeometry(40, 40),
		new THREE.ShadowMaterial({ opacity: 0.22 }),
	);
	ground.rotation.x = -Math.PI / 2;
	ground.receiveShadow = true;
	scene.add(ground);

	const c = buildChungus();
	scene.add(c.root);
	c.root.position.x = 9;
	c.root.rotation.y = -Math.PI / 2;
	c.arm.pivot.quaternion.copy(ARM_REST);

	const ring = new THREE.Mesh(
		new THREE.RingGeometry(0.3, 0.38, 64),
		new THREE.MeshBasicMaterial({
			color: 0xffd36b,
			transparent: true,
			opacity: 0,
			side: THREE.DoubleSide,
			// Drawn over everything so the ground never clips the shockwave.
			depthTest: false,
		}),
	);
	ring.renderOrder = 1;
	scene.add(ring);

	// Sparks thrown off where the thumb and middle finger part.
	const sparkMaterial = new THREE.MeshBasicMaterial({
		color: 0xffe08a,
		transparent: true,
		toneMapped: false,
	});
	const sparks = Array.from({ length: 28 }, () => {
		const spark = new THREE.Mesh(new THREE.SphereGeometry(0.025, 6, 4), sparkMaterial);
		spark.visible = false;
		spark.userData.velocity = new THREE.Vector3();
		scene.add(spark);
		return spark;
	});

	const camera = new THREE.PerspectiveCamera(30, 1, 0.1, 100);
	const view = { base: new THREE.Vector3(), target: new THREE.Vector3(), shake: 0 };
	const hand = new THREE.Vector3();
	const state = {
		walking: false,
		charge: 0,
		idle: false,
		pose: { from: HANDS.fist, to: HANDS.fist, t: 1 },
	};

	function layout() {
		const w = host.clientWidth;
		const h = host.clientHeight;
		renderer.setSize(w, h);
		camera.aspect = w / h;
		const narrow = w < 760;
		view.target.set(0, narrow ? 3.7 : 2.35, 0);
		view.base.set(0, narrow ? 3.4 : 2.7, narrow ? 17.5 : 14);
		camera.updateProjectionMatrix();
	}
	const resize = new ResizeObserver(layout);
	resize.observe(host);
	layout();

	/** Tweens run inside the render loop. */
	const tweens = new Set();
	const tween = (ms, fn) =>
		new Promise((resolve) =>
			tweens.add({ start: performance.now(), ms: fast ? ms / 3 : ms, fn, resolve }),
		);

	function setFingers(from, to, t) {
		const lerp = (x, y) => THREE.MathUtils.lerp(x, y, t);
		c.gauntlet.fingers.forEach((joints, i) =>
			joints.forEach((j, k) => (j.rotation.x = lerp(from.fingers[i][k], to.fingers[i][k]))),
		);
		const [base, tip] = c.gauntlet.thumb;
		base.rotation.set(
			lerp(from.thumb[0], to.thumb[0]),
			lerp(from.thumb[1], to.thumb[1]),
			lerp(from.thumb[2], to.thumb[2]),
		);
		tip.rotation.x = lerp(from.thumb[3], to.thumb[3]);
	}
	setFingers(HANDS.fist, HANDS.fist, 1);
	const shapeHand = (name, ms, curve = ease) => {
		const from = state.pose.to;
		state.pose.to = HANDS[name];
		return tween(ms, (t) => setFingers(from, HANDS[name], curve(t)));
	};

	let frame = 0;
	let lastStep = 0;
	const clock = new THREE.Clock();
	function render() {
		frame = requestAnimationFrame(render);
		const now = performance.now();
		const time = clock.getElapsedTime();
		for (const tw of tweens) {
			const t = Math.min(1, (now - tw.start) / tw.ms);
			tw.fn(t);
			if (t === 1) {
				tweens.delete(tw);
				tw.resolve();
			}
		}

		if (state.walking) {
			const phase = time * (Math.PI / 0.42);
			c.figure.rotation.z = Math.sin(phase) * 0.07;
			c.figure.position.y = Math.abs(Math.sin(phase)) * 0.14;
			c.feet[0].position.y = 0.13 + Math.max(0, Math.sin(phase)) * 0.22;
			c.feet[1].position.y = 0.13 + Math.max(0, -Math.sin(phase)) * 0.22;
			const step = Math.floor(phase / Math.PI);
			if (step !== lastStep) {
				lastStep = step;
				view.shake = 0.08;
				onStep();
			}
		} else if (state.idle) {
			c.figure.scale.y = 1 + Math.sin(time * 2) * 0.012;
			c.ears[1].rotation.x = Math.max(0, Math.sin(time * 1.3)) * -0.35;
		}

		const g = c.gauntlet;
		const pulse = state.charge * (0.7 + Math.sin(time * 22) * 0.25);
		for (const stone of g.stones) stone.emissiveIntensity = 0.25 + pulse;
		g.glow.intensity = state.charge * 4;

		g.hand.localToWorld(hand.set(0, -0.7, 0));
		// The camera holds still; only footsteps nudge it.
		view.shake = Math.max(0, view.shake * 0.88 - 0.0005);
		camera.position.copy(view.base);
		camera.position.x += (Math.random() - 0.5) * view.shake;
		camera.position.y += (Math.random() - 0.5) * view.shake;
		camera.lookAt(view.target);
		renderer.render(scene, camera);
	}
	render();

	const toScreen = (point) => {
		const p = point.clone().project(camera);
		return { x: ((p.x + 1) / 2) * host.clientWidth, y: ((1 - p.y) / 2) * host.clientHeight };
	};

	return {
		async walkIn() {
			state.walking = true;
			await tween(3000, (t) => (c.root.position.x = 9 * (1 - ease(Math.min(1, t * 1.05)))));
			state.walking = false;
			c.figure.rotation.z = 0;
			c.figure.position.y = 0;
			for (const foot of c.feet) foot.position.y = 0.13;
		},
		face: () =>
			tween(650, (t) => {
				c.root.rotation.y = (-Math.PI / 2) * (1 - easeOutBack(t));
				c.figure.position.y = Math.sin(t * Math.PI) * 0.25;
			}),
		raise: () =>
			Promise.all([
				tween(900, (t) => {
					c.arm.pivot.quaternion.slerpQuaternions(ARM_REST, ARM_UP, easeOutBack(t));
					// Turn the thumb side to the camera so the snap reads like a real one.
					c.arm.wrist.rotation.set(WRIST_UP.x * ease(t), WRIST_UP.y * ease(t), 0);
					// The gauntlet swells a little as it powers up, so the fingers read from afar.
					c.gauntlet.hand.scale.setScalar(1.55 + 0.45 * ease(t));
				}),
				shapeHand('ready', 900),
			]),
		charge: () => tween(1400, (t) => (state.charge = ease(t))),
		async snap() {
			await shapeHand('press', 260);
			const tip = c.gauntlet.fingers[1][2].localToWorld(new THREE.Vector3(0, -0.12, -0.05));
			for (const spark of sparks) {
				spark.userData.velocity.randomDirection().multiplyScalar(0.8 + Math.random() * 1.2);
				spark.visible = true;
			}
			// Sparks drift out on real time, so the burst is smooth at any frame rate.
			void tween(900, (t) => {
				sparkMaterial.opacity = 1 - easeOut(t);
				for (const spark of sparks) {
					spark.position.copy(tip).addScaledVector(spark.userData.velocity, easeOut(t));
					spark.visible = t < 1;
				}
			});
			// A soft follow-through of the wrist as the middle finger leaves the thumb.
			void tween(
				700,
				(t) =>
					(c.arm.wrist.rotation.x = WRIST_UP.x + Math.sin(t * Math.PI) * easeOut(1 - t) * 0.18),
			);
			await shapeHand('snapped', 150, easeOut);
			ring.position.copy(hand);
			ring.lookAt(camera.position);
			void tween(900, (t) => {
				ring.scale.setScalar(1 + t * 14);
				ring.material.opacity = (1 - t) * 0.9;
			});
		},
		release: () => tween(1100, (t) => (state.charge = 1 - ease(t))),
		async relax() {
			await Promise.all([
				tween(900, (t) => {
					c.arm.pivot.quaternion.slerpQuaternions(ARM_UP, ARM_REST, ease(t));
					c.arm.wrist.rotation.set(WRIST_UP.x * (1 - ease(t)), WRIST_UP.y * (1 - ease(t)), 0);
					c.gauntlet.hand.scale.setScalar(2 - 0.45 * ease(t));
				}),
				shapeHand('fist', 900),
			]);
			state.idle = true;
		},
		// Beside the head on wide screens, above it on narrow ones (see chungus.css).
		headPoint: () =>
			toScreen(
				c.head
					.getWorldPosition(new THREE.Vector3())
					.add(
						host.clientWidth < 760
							? new THREE.Vector3(0, 0.75, 0)
							: new THREE.Vector3(-0.65, 0.05, 0),
					),
			),
		handPoint: () => toScreen(hand),
		dispose() {
			cancelAnimationFrame(frame);
			resize.disconnect();
			for (const tw of tweens) tw.resolve();
			tweens.clear();
			pmrem.dispose();
			renderer.dispose();
			scene.traverse((o) => {
				if (o instanceof THREE.Mesh) {
					o.geometry.dispose();
					/** @type {THREE.Material} */ (o.material).dispose();
				}
			});
		},
	};
}
