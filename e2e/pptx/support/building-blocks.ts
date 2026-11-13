/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright support */
import { test } from '@playwright/test';

/**
 * Building-block specs run for the bindings whose demo has the `?blocks=1` page: their
 * Playwright project carries `metadata.buildingBlocks` (`viewers/pptx/playwright.config.ts`).
 * A port adds its project there once its demo page exists; the spec itself never changes.
 */
export function requireBuildingBlocks(): void {
	test.skip(
		!test.info().project.metadata.buildingBlocks,
		'This binding has no building-blocks demo page yet.',
	);
}
