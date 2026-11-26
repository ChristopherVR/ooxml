import {
	loadVisioStencilEditor,
	readVisioStencil,
	type VisioStencilEdit,
	type VisioStencilFile,
	type VisioStencilSave,
} from 'ooxml-core/visio';

export type StencilWorkerRequest =
	| { type: 'read'; source: Uint8Array }
	| { type: 'create' }
	| { type: 'edit'; source: Uint8Array; edit?: VisioStencilEdit };

export type StencilWorkerResponse =
	| { type: 'read'; value: VisioStencilFile }
	| { type: 'create'; value: Uint8Array }
	| { type: 'edit'; value: VisioStencilSave }
	| { type: 'error'; message: string };

export async function processStencilRequest(
	request: StencilWorkerRequest,
): Promise<Exclude<StencilWorkerResponse, { type: 'error' }>> {
	if (request.type === 'read')
		return { type: 'read', value: await readVisioStencil(request.source) };
	const editor = await loadVisioStencilEditor();
	if (request.type === 'create')
		return { type: 'create', value: await editor.createVisioStencil() };
	return { type: 'edit', value: await editor.editVisioStencil(request.source, request.edit) };
}
