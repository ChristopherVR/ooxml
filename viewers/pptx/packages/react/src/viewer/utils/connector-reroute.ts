/**
 * Thin re-export shim → `pptx-viewer-shared`.
 *
 * Connector dynamic rerouting (recalculating connector endpoints when connected
 * shapes move/resize) now lives in `pptx-viewer-shared`
 * (`render/connector-reroute`). This shim preserves the historical import
 * surface so the editor hooks and colocated tests are unchanged.
 */
export type { ReroutedConnector } from 'ooxml-ui/pptx';
export {
	rerouteConnectorsForMovedElements,
	computeConnectorGeometry,
	applyReroutedConnectors,
} from 'ooxml-ui/pptx';
