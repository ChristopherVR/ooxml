/** Namespace URIs used across the OOXML part families, so parsers never repeat the literals. */
export const NS = {
	/** WordprocessingML main. */
	w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
	/** PresentationML main. */
	p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
	/** SpreadsheetML main. */
	x: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
	/** DrawingML main. */
	a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
	/** DrawingML chart. */
	c: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
	/** DrawingML diagram (SmartArt) main. */
	dgm: 'http://schemas.openxmlformats.org/drawingml/2006/diagram',
	/** Office 2007+ diagram drawing (`dsp`). */
	dsp: 'http://schemas.microsoft.com/office/drawing/2008/diagram',
	/** DrawingML Word drawing (`wp:inline`, `wp:anchor`). */
	wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
	/** DrawingML picture. */
	pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
	/** Word 2010 shapes (`wps`). */
	wps: 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
	/** Relationships (`r:id`, `r:embed`). */
	r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
	/** Markup compatibility (`mc:AlternateContent`). */
	mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
	/** OPC relationships part. */
	rels: 'http://schemas.openxmlformats.org/package/2006/relationships',
	/** OPC content types part. */
	contentTypes: 'http://schemas.openxmlformats.org/package/2006/content-types',
	/** Office Math. */
	m: 'http://schemas.openxmlformats.org/officeDocument/2006/math',
	/** Core properties. */
	cp: 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
	/** Dublin Core elements. */
	dc: 'http://purl.org/dc/elements/1.1/',
	/** Legacy VML. */
	v: 'urn:schemas-microsoft-com:vml',
	/** Office VML extensions. */
	o: 'urn:schemas-microsoft-com:office:office',
	/** Word VML extensions. */
	w10: 'urn:schemas-microsoft-com:office:word',
	/** SpreadsheetML drawing (`xdr:wsDr`, cell anchors). */
	xdr: 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
	/** SpreadsheetML 2009 extensions (`x14:conditionalFormatting`, `x14:dataBar`). */
	x14: 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main',
	/** SpreadsheetML 2009 attribute extensions (`x14ac:dyDescent`). */
	x14ac: 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac',
	/** SpreadsheetML revision ids (`xr:uid`). */
	xr: 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision',
	/** Excel threaded comments and persons. */
	tc: 'http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments',
	/** Legacy VML Excel client data (`x:ClientData`). */
	xv: 'urn:schemas-microsoft-com:office:excel',
	/** Extended (app) properties. */
	extendedProperties: 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
	/** Dublin Core terms (`dcterms:created`). */
	dcterms: 'http://purl.org/dc/terms/',
	/** XML itself (`xml:space`). */
	xml: 'http://www.w3.org/XML/1998/namespace',
} as const;
export type NamespaceKey = keyof typeof NS;
