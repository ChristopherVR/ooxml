/**
 * Context labels of the coefficient bit modelling (ITU-T T.800 D.3): zero coding (Table D.1),
 * sign coding (Table D.3) and magnitude refinement (Table D.4), with the initial states (D.7).
 */

export const RUN_LENGTH = 17;
export const UNIFORM = 18;

/** Zero coding labels by orientation (0 LL, 1 HL, 2 LH, 3 HH), indexed h * 15 + v * 5 + d. */
export const ZERO_CODING: Uint8Array[] = [0, 1, 2, 3].map((orientation) => {
	const table = new Uint8Array(45);
	for (let h = 0; h <= 2; h++)
		for (let v = 0; v <= 2; v++)
			for (let d = 0; d <= 4; d++) {
				let label: number;
				if (orientation === 3) {
					const hv = h + v;
					if (d >= 3) label = 8;
					else if (d === 2) label = hv >= 1 ? 7 : 6;
					else if (d === 1) label = hv >= 2 ? 5 : hv === 1 ? 4 : 3;
					else label = hv >= 2 ? 2 : hv;
				} else {
					// HL (horizontally high-pass) swaps the roles of H and V.
					const [a, b] = orientation === 1 ? [v, h] : [h, v];
					if (a === 2) label = 8;
					else if (a === 1) label = b >= 1 ? 7 : d >= 1 ? 6 : 5;
					else if (b === 2) label = 4;
					else if (b === 1) label = 3;
					else label = d >= 2 ? 2 : d;
				}
				table[h * 15 + v * 5 + d] = label;
			}
	return table;
});

/** A fresh set of 19 context states (index << 1 | MPS). */
export function initialContexts(contexts: Uint8Array): void {
	contexts.fill(0);
	contexts[0] = 4 << 1;
	contexts[RUN_LENGTH] = 3 << 1;
	contexts[UNIFORM] = 46 << 1;
}
