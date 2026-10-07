// Compatibility exports: document operations live in ooxml-core.
export {
	YDOC_ASSETS_KEY,
	ASSET_ELEMENT_FIELDS,
	assetVersionKey,
	isAssetVersionKey,
	assetKey,
	getAssetsMap,
	isAssetRefKey,
	writeAssetFields,
	reconcileAssetFields,
	readAssetFields,
} from 'ooxml-core/pptx/editor/render/collaboration-assets';
