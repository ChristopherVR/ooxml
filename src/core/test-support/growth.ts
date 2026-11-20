// Test-only: how a run's time grows with its input, for the linear-time checks. Measured so that
// machine load does not decide the answer: the two sizes take turns, so a busy spell slows both,
// and each keeps the fastest of its runs. Four times the input reads about 4 when linear and 16
// when quadratic, so the checks ask for less than 8.

const ROUNDS = 5;

/** The fastest time of `large` over the fastest time of `small`, measured alternately. */
export function growth(small: () => unknown, large: () => unknown, rounds = ROUNDS): number {
	let fastSmall = Number.POSITIVE_INFINITY;
	let fastLarge = Number.POSITIVE_INFINITY;
	for (let round = 0; round < rounds; round++) {
		let start = performance.now();
		small();
		fastSmall = Math.min(fastSmall, performance.now() - start);
		start = performance.now();
		large();
		fastLarge = Math.min(fastLarge, performance.now() - start);
	}
	return fastLarge / fastSmall;
}

/** `growth` for asynchronous runs. */
export async function growthAsync(
	small: () => Promise<unknown>,
	large: () => Promise<unknown>,
	rounds = ROUNDS,
): Promise<number> {
	let fastSmall = Number.POSITIVE_INFINITY;
	let fastLarge = Number.POSITIVE_INFINITY;
	for (let round = 0; round < rounds; round++) {
		let start = performance.now();
		await small();
		fastSmall = Math.min(fastSmall, performance.now() - start);
		start = performance.now();
		await large();
		fastLarge = Math.min(fastLarge, performance.now() - start);
	}
	return fastLarge / fastSmall;
}
