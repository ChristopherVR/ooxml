import { getTestBed } from '@angular/core/testing';
import { afterAll } from 'vitest';

// vi.resetModules() drops the fixture's Vite modules. Angular's native TestBed singleton stays
// loaded, so release the previous testing environment before the next file initializes it.
afterAll(() => {
	const testBed = getTestBed();
	if (testBed.platform) {
		testBed.resetTestEnvironment();
	}
});
