import { guidFrom, randomIdSource, type IdSource } from '../opc/reproducible';

/**
 * A fresh braced upper-case GUID for an Office rule, series or part id (never an encryption
 * secret), drawn from `ids`: random by default, a seeded source for a reproducible save.
 */
export function createOfficeGuid(ids: IdSource = randomIdSource): string {
	return guidFrom(ids);
}
