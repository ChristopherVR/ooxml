// @vitest-environment jsdom
import { afterEach, expect, it, vi } from 'vitest';
import { definePptxCameoPreview } from './cameo-preview';

definePptxCameoPreview(customElements);
afterEach(() => {
	document.body.replaceChildren();
	vi.unstubAllGlobals();
});

function mount() {
	const cameo = document.createElement('pptx-ui-cameo-preview');
	document.body.append(cameo);
	return { cameo, button: cameo.shadowRoot!.querySelector('button')! };
}

it('requests video only on activation and stops the track when leaving the slide', async () => {
	const stop = vi.fn();
	const stream = { getTracks: () => [{ stop }] };
	const getUserMedia = vi.fn().mockResolvedValue(stream);
	vi.stubGlobal('navigator', { mediaDevices: { getUserMedia } });
	const { cameo, button } = mount();
	expect(getUserMedia).not.toHaveBeenCalled();
	button.click();
	await vi.waitFor(() => {
		expect(cameo.shadowRoot?.querySelector('video')?.srcObject).toBe(stream);
	});
	expect(getUserMedia).toHaveBeenCalledExactlyOnceWith({ video: true, audio: false });
	cameo.remove();
	expect(stop).toHaveBeenCalledOnce();
});

it('stops a late stream immediately if the preview has left the document', async () => {
	const stop = vi.fn();
	let grant!: (stream: { getTracks: () => { stop: typeof stop }[] }) => void;
	const pending = new Promise<Parameters<typeof grant>[0]>((resolve) => (grant = resolve));
	vi.stubGlobal('navigator', { mediaDevices: { getUserMedia: () => pending } });
	const { cameo, button } = mount();
	button.click();
	cameo.remove();
	grant({ getTracks: () => [{ stop }] });
	await vi.waitFor(() => expect(stop).toHaveBeenCalledOnce());
});
