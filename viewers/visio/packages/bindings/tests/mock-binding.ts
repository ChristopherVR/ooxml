import { vi } from 'vitest';
import type { ViewerOptions, ViewerEvents } from '../../../src/contract.js';
import type { MountedViewer } from '../../../src/binding.js';
export interface FakeBinding {
	binding: MountedViewer;
	options: ViewerOptions;
	destroyed: boolean;
	update: ReturnType<typeof vi.fn>;
	destroy: ReturnType<typeof vi.fn>;
}
const mocks = vi.hoisted(() => ({ instances: [] as FakeBinding[], mount: vi.fn() }));
vi.mock('../../../src/binding.js', () => ({
	mountViewer: mocks.mount.mockImplementation((host: HTMLElement, initial: ViewerOptions) => {
		const element = document.createElement('section');
		element.dataset.testViewer = '';
		host.append(element);
		const instance: FakeBinding = {} as FakeBinding;
		const assertAlive = () => {
			if (instance.destroyed) throw new Error('The viewer has been destroyed.');
		};
		const update = vi.fn((options: ViewerOptions) => {
			assertAlive();
			instance.options = { ...instance.options, ...options };
		});
		const destroy = vi.fn(() => {
			instance.destroyed = true;
			element.remove();
		});
		Object.assign(instance, {
			options: initial,
			destroyed: false,
			update,
			destroy,
			binding: {
				element,
				controller: { marker: 'controller' },
				update,
				destroy,
				load: vi.fn(async () => {
					assertAlive();
				}),
				fit: vi.fn(assertAlive),
				exportSvg: vi.fn(() => {
					assertAlive();
					return {
						svg: '<svg/>',
						byteLength: 6,
						pageIndex: 0,
						pageName: 'Mock',
						width: 1,
						height: 1,
						diagnostics: [],
					};
				}),
			} as unknown as MountedViewer,
		});
		mocks.instances.push(instance);
		return instance.binding;
	}),
}));
vi.mock('../../../src/controller.js', () => ({ ViewerController: class ViewerController {} }));
export function current(): FakeBinding {
	return mocks.instances.at(-1)!;
}
export function emit<K extends keyof ViewerEvents>(name: K, value: ViewerEvents[K]): void {
	current().options.events?.[name]?.(value);
}
export function reset(): void {
	mocks.instances.length = 0;
	mocks.mount.mockClear();
}

export { mocks };
