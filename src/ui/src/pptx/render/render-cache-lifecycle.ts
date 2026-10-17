import { clearBarFacePicturePixelCache } from './chart-bar3d-face-picture-sample';
import { clearColorChangeCache } from './image-color-change';
import { clearDuotoneCache } from './image-duotone-canvas';
import { clearNativeImageSizeCache } from './image-native-size';
import { invalidatePptxRenderCacheEpoch } from './render-cache';

/**
 * Release module-level image memoisation after a host closes all its documents.
 * This is global to this module instance. Call after disposing every owned
 * stage, when no other viewer sharing this instance needs pending samples.
 * Existing async operations may still finish for their caller;
 * their results do not repopulate these caches or notify chart subscribers.
 * Does not revoke host-owned Blob URLs, detach stages or purge browser caches.
 */
export function releasePptxRenderCaches(): void {
	invalidatePptxRenderCacheEpoch();
	clearNativeImageSizeCache();
	clearBarFacePicturePixelCache();
	clearColorChangeCache();
	clearDuotoneCache();
}
