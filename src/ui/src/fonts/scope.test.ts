import { afterEach, expect, expectTypeOf, it, vi } from 'vitest';
import {
	createFontFaceScope,
	type CustomFontRegistration,
	type CustomFontDescriptor,
	type OwnedCustomFontRegistration,
} from './custom-fonts';

afterEach(() => vi.unstubAllGlobals());

it('retains the original structural registration API', () => {
	expectTypeOf<{
		family: string;
		descriptor: CustomFontDescriptor;
	}>().toExtend<CustomFontRegistration>();
	expectTypeOf<OwnedCustomFontRegistration>().toExtend<CustomFontRegistration>();
});

it('removes only owned faces, including a load completed after disposal', async () => {
	let finish!: () => void;
	const ready = new Promise<void>((resolve) => {
		finish = resolve;
	});
	class Face {
		load = () => ready.then(() => this);
	}
	vi.stubGlobal('FontFace', Face);
	const add = vi.fn();
	const remove = vi.fn();
	const scope = createFontFaceScope({ add, delete: remove } as unknown as FontFaceSet);
	const file = { name: 'Acme-Bold.ttf', arrayBuffer: async () => new ArrayBuffer(1) } as File;
	const pending = scope.register(file);
	await Promise.resolve();
	scope.dispose();
	finish();
	expect(await pending).toBeNull();
	expect(add).toHaveBeenCalledOnce();
	expect(remove).toHaveBeenCalledWith(add.mock.calls[0]![0]);
	expect(await scope.register(file)).toBeNull();
	scope.dispose();
	expect(remove).toHaveBeenCalledOnce();
});

it('does not register a failed font', async () => {
	class Face {
		load = async () => {
			throw new Error('invalid font');
		};
	}
	vi.stubGlobal('FontFace', Face);
	const add = vi.fn();
	const scope = createFontFaceScope({ add } as unknown as FontFaceSet);
	await expect(
		scope.register({ name: 'Acme.ttf', arrayBuffer: async () => new ArrayBuffer(1) } as File),
	).rejects.toThrow('invalid font');
	expect(add).not.toHaveBeenCalled();
	scope.dispose();
});
