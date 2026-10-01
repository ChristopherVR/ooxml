const BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** Relationship type URIs for the parts the Office viewers read and write. */
export const RELATIONSHIP_TYPES = {
	officeDocument: `${BASE}/officeDocument`,
	styles: `${BASE}/styles`,
	numbering: `${BASE}/numbering`,
	settings: `${BASE}/settings`,
	theme: `${BASE}/theme`,
	fontTable: `${BASE}/fontTable`,
	webSettings: `${BASE}/webSettings`,
	header: `${BASE}/header`,
	footer: `${BASE}/footer`,
	footnotes: `${BASE}/footnotes`,
	endnotes: `${BASE}/endnotes`,
	comments: `${BASE}/comments`,
	image: `${BASE}/image`,
	hyperlink: `${BASE}/hyperlink`,
	oleObject: `${BASE}/oleObject`,
	chart: `${BASE}/chart`,
	package: `${BASE}/package`,
	slide: `${BASE}/slide`,
	slideLayout: `${BASE}/slideLayout`,
	slideMaster: `${BASE}/slideMaster`,
	notesSlide: `${BASE}/notesSlide`,
	diagramData: `${BASE}/diagramData`,
	diagramLayout: `${BASE}/diagramLayout`,
	diagramQuickStyle: `${BASE}/diagramQuickStyle`,
	diagramColors: `${BASE}/diagramColors`,
	diagramDrawing: 'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing',
	coreProperties:
		'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
	extendedProperties: `${BASE}/extended-properties`,
	customXml: `${BASE}/customXml`,
} as const;
export type RelationshipTypeKey = keyof typeof RELATIONSHIP_TYPES;
