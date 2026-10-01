// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Relationship parsing and part-path resolution live in the sibling `opc` area;
// this module keeps the existing import path working.
export {
	getRelationshipId,
	parseRelationships,
	resolvePartPath,
	type Relationship,
} from '../opc/index.js';
import { NS } from '../xml/index.js';

export const RELATIONSHIP_NS = NS.r;
