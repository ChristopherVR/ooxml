/**
 * Insert/Duplicate/Delete/Rename Slide Master. Split out of
 * {@link module:sdk/master-layout-crud} purely for file size; see that
 * module's header for the overall strategy (save -> ZIP surgery -> reload).
 *
 * @module sdk/master-layout-crud-master
 */
import type JSZip from 'jszip';
import { clonePart, clonePartGraph, relativePartTarget } from '../../../../opc/clone-part-graph';
import type { Relationship } from '../../../../opc/relationships';
import { copiesRelationshipTarget } from '../../core/builders/duplicate-slide-parts';
import { PptxHandler } from '../../PptxHandler';
import type { PptxData } from '../../types/presentation';
import {
	addLayoutContentType,
	addLayoutToSlideMaster,
	countExistingLayouts,
	generateLayoutXml,
	layoutRelsXml,
} from './layout-operations';
import { findMaster, NOT_FOUND } from './master-layout-crud-lookup';
import { collectMasterNames, uniquePrefixedName } from './master-layout-crud-names';
import {
	addMasterContentType,
	addMasterToPresentation,
	countExistingMasters,
	newMasterXmlFromSource,
	relsPathFor,
	reload,
	removeContentTypeOverride,
	removeMasterFromPresentation,
	renumberMasterLayoutIds,
	saveToZip,
	withCSldName,
} from './master-layout-crud-xml';
import type { MasterLayoutCrudFailure, MasterLayoutCrudResult } from './master-layout-crud-xml';
import { STANDARD_LAYOUTS } from './PresentationBuilder';
import { at } from '../../../../geometry/indexed';

/** {@link duplicateSlideMaster} / {@link insertSlideMaster} success: the new master's id. */
export interface DuplicateMasterSuccess {
	ok: true;
	handler: PptxHandler;
	data: PptxData;
	masterId: string;
}
export type DuplicateMasterResult = DuplicateMasterSuccess | MasterLayoutCrudFailure;

/** A master owns its layouts and its theme, so a copy of it gets its own of each. */
const copiesWithMaster = (relationship: Relationship): boolean =>
	copiesRelationshipTarget(relationship) || /\/(?:slideLayout|theme)$/u.test(relationship.type);

/** The first `ppt/slideMasters/slideMaster<n>.xml` the package does not hold. */
function freeMasterPath(zip: JSZip): string {
	let index = 1;
	while (zip.file(`ppt/slideMasters/slideMaster${index}.xml`)) index += 1;
	return `ppt/slideMasters/slideMaster${index}.xml`;
}

/**
 * Deep-copy a slide master together with every layout it owns and its theme, naming it with
 * PowerPoint's "duplicate master" convention ({@link uniquePrefixedName}). The copy owns new
 * layout, theme and element parts with fresh names; PowerPoint refuses a package where two
 * masters share one theme part.
 */
export async function duplicateSlideMaster(
	handler: PptxHandler,
	data: PptxData,
	masterId: string,
): Promise<DuplicateMasterResult> {
	const source = findMaster(data, masterId);
	if (!source) {
		return NOT_FOUND;
	}

	const zip = await saveToZip(handler, data);
	const sourceXml = await zip.file(masterId)?.async('string');
	if (!sourceXml) {
		return NOT_FOUND;
	}

	const newMasterPath = freeMasterPath(zip);
	const newName = uniquePrefixedName(collectMasterNames(data), source.name ?? 'Office Theme');
	zip.file(newMasterPath, withCSldName(sourceXml, newName));
	await clonePartGraph({
		zip,
		sourcePart: masterId,
		targetPart: newMasterPath,
		shouldClone: copiesWithMaster,
	});
	await renumberMasterLayoutIds(zip, newMasterPath);
	await addMasterToPresentation(zip, newMasterPath);

	const result = await reload(zip);
	return { ...result, masterId: newMasterPath };
}

/**
 * Remove a slide master and every layout it owns. Refuses when it is the
 * presentation's only master (`lastMaster`) or a slide still uses one of its
 * layouts (`inUse`).
 */
export async function deleteSlideMaster(
	handler: PptxHandler,
	data: PptxData,
	masterId: string,
): Promise<MasterLayoutCrudResult> {
	const masters = data.slideMasters ?? [];
	if (masters.length <= 1) {
		return { ok: false, reason: 'lastMaster' };
	}
	const master = masters.find((m) => m.path === masterId);
	if (!master) {
		return NOT_FOUND;
	}
	const layoutPaths = new Set((master.layouts ?? []).map((l) => l.path));
	const inUse = data.slides.some(
		(slide) => slide.layoutPath !== undefined && layoutPaths.has(slide.layoutPath),
	);
	if (inUse) {
		return { ok: false, reason: 'inUse' };
	}

	const zip = await saveToZip(handler, data);
	await removeMasterFromPresentation(zip, masterId);
	await removeContentTypeOverride(zip, masterId);
	zip.remove(masterId);
	zip.remove(relsPathFor(masterId));
	for (const layoutPath of layoutPaths) {
		await removeContentTypeOverride(zip, layoutPath);
		zip.remove(layoutPath);
		zip.remove(relsPathFor(layoutPath));
	}

	return reload(zip);
}

/** Rename a slide master (`p:cSld/@name`) via the existing typed-model save path. */
export async function renameSlideMaster(
	handler: PptxHandler,
	data: PptxData,
	masterId: string,
	name: string,
): Promise<MasterLayoutCrudResult> {
	const master = findMaster(data, masterId);
	if (!master) {
		return NOT_FOUND;
	}
	const slideMasters = (data.slideMasters ?? []).map((m) =>
		m.path === masterId ? { ...m, name } : m,
	);
	const bytes = await handler.save(data.slides, { slideMasters });
	const newHandler = new PptxHandler();
	const newData = await newHandler.load(bytes.buffer as ArrayBuffer);
	return { ok: true, handler: newHandler, data: newData };
}

/**
 * Insert a brand-new slide master: PowerPoint's own eleven standard layouts
 * ({@link STANDARD_LAYOUTS}, the same set `PresentationBuilder` seeds a
 * from-scratch deck with), the first master's colour map and text styles
 * byte-copied across, with its own copy of that master's theme part.
 */
export async function insertSlideMaster(
	handler: PptxHandler,
	data: PptxData,
): Promise<DuplicateMasterResult> {
	const firstMaster = data.slideMasters?.[0];
	if (!firstMaster) {
		return NOT_FOUND;
	}

	const zip = await saveToZip(handler, data);
	const sourceMasterXml = await zip.file(firstMaster.path)?.async('string');
	if (!sourceMasterXml) {
		return NOT_FOUND;
	}

	const masterIndex = countExistingMasters(zip) + 1;
	const newMasterPath = `ppt/slideMasters/slideMaster${masterIndex}.xml`;
	// The new master owns a copy of the first master's theme: two masters naming one theme part
	// make PowerPoint refuse the package.
	const sourceTheme = firstMaster.themePath ?? 'ppt/theme/theme1.xml';
	if (!zip.file(sourceTheme)) {
		return NOT_FOUND;
	}
	const themeCopy = (
		await clonePart({ zip, part: sourceTheme, shouldClone: copiesRelationshipTarget })
	).get(sourceTheme)!;
	const themeTarget = relativePartTarget(newMasterPath, themeCopy);

	zip.file(newMasterPath, newMasterXmlFromSource(sourceMasterXml, ''));
	zip.file(
		relsPathFor(newMasterPath),
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="${themeTarget}"/>\n</Relationships>`,
	);
	await addMasterContentType(zip, masterIndex);

	const layoutBase = countExistingLayouts(zip);
	for (let i = 0; i < STANDARD_LAYOUTS.length; i++) {
		const newIndex = layoutBase + i + 1;
		const layoutPath = `ppt/slideLayouts/slideLayout${newIndex}.xml`;
		zip.file(layoutPath, generateLayoutXml(at(STANDARD_LAYOUTS, i)));
		zip.file(relsPathFor(layoutPath), layoutRelsXml(masterIndex));
		await addLayoutToSlideMaster(zip, newIndex, masterIndex);
		await addLayoutContentType(zip, newIndex);
	}

	await addMasterToPresentation(zip, newMasterPath);

	const result = await reload(zip);
	return { ...result, masterId: newMasterPath };
}
