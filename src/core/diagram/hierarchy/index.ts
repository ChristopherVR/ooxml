// The layout interpreter's hierarchy arranger (`hierRoot`/`hierChild` org charts, hierarchy lists,
// hanging and fanned trees): places a node forest from the layout definition's declared
// constraints. `arrangeHierarchy` takes the `RawXmlView` of the layout definition's raw slots.
export { arrangeHierarchy } from './smartart-layout-interpreter-hierarchy';
export {
	resolveHierarchyDispatchChAlign,
	resolveHierarchyDispatchLinDir,
	resolveHierarchyRootAlign,
} from './smartart-hierarchy-dispatch-lindir';
export {
	resolveHierarchyGenerationTemplates,
	type HierarchyGenerationTemplate,
	type HierarchyGenerationTemplates,
	type HierarchyRootTemplate,
} from './smartart-hierarchy-generation-templates';
export { tailedHierarchyDeclaresChAlign } from './smartart-hierarchy-tailed-transpose';
