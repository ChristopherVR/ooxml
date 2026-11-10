// @vitest-environment happy-dom
import type { MasterViewCrudAction } from 'ooxml-ui/pptx';
import React, { act } from 'react';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { createMarkedRoot, marked } from '../i18n/marking-translator.test-support';
import { MasterViewSidebar } from './MasterViewSidebar';

let container: HTMLDivElement;
let root: Root;

beforeEach(() => {
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createMarkedRoot(container);
});

afterEach(() => {
	act(() => root.unmount());
	container.remove();
});

const CRUD_ACTIONS: MasterViewCrudAction[] = [
	{ id: 'addLayout', labelKey: 'pptx.masterView.addLayout', enabled: true },
	{
		id: 'deleteLayout',
		labelKey: 'pptx.masterView.deleteLayout',
		enabled: false,
		disabledReasonKey: 'pptx.masterView.layoutInUse',
	},
];

describe('masterViewSidebar CRUD buttons', () => {
	it('renders one button per action, disabled ones carry the reason as a title', () => {
		act(() => {
			root.render(
				<MasterViewSidebar
					slideMasters={[]}
					activeMasterIndex={0}
					activeLayoutIndex={null}
					canvasSize={{ width: 960, height: 540 }}
					masterViewTab='slides'
					notesMaster={undefined}
					handoutMaster={undefined}
					handoutSlidesPerPage={6}
					canEdit
					onSelectMaster={() => {}}
					onSelectLayout={() => {}}
					onCollapse={() => {}}
					onTabChange={() => {}}
					crudActions={CRUD_ACTIONS}
					onCrudAction={() => {}}
					onHandoutSlidesPerPageChange={() => {}}
					onNotesMasterBackgroundChange={() => {}}
					onHandoutMasterBackgroundChange={() => {}}
					onSlidesBackgroundChange={() => {}}
				/>,
			);
		});

		const addButton = container.querySelector(
			'[data-testid="pptx-master-crud-addLayout"]',
		) as HTMLButtonElement;
		const deleteButton = container.querySelector(
			'[data-testid="pptx-master-crud-deleteLayout"]',
		) as HTMLButtonElement;
		expect(addButton).not.toBeNull();
		expect(addButton.disabled).toBeFalsy();
		expect(deleteButton.disabled).toBeTruthy();
		// The marking translator shows which key reached the title attribute.
		expect(deleteButton.title).toBe(marked('pptx.masterView.layoutInUse'));
	});

	it('clicking an enabled button calls onCrudAction with its id', () => {
		const onCrudAction = vi.fn();
		act(() => {
			root.render(
				<MasterViewSidebar
					slideMasters={[]}
					activeMasterIndex={0}
					activeLayoutIndex={null}
					canvasSize={{ width: 960, height: 540 }}
					masterViewTab='slides'
					notesMaster={undefined}
					handoutMaster={undefined}
					handoutSlidesPerPage={6}
					canEdit
					onSelectMaster={() => {}}
					onSelectLayout={() => {}}
					onCollapse={() => {}}
					onTabChange={() => {}}
					crudActions={CRUD_ACTIONS}
					onCrudAction={onCrudAction}
					onHandoutSlidesPerPageChange={() => {}}
					onNotesMasterBackgroundChange={() => {}}
					onHandoutMasterBackgroundChange={() => {}}
					onSlidesBackgroundChange={() => {}}
				/>,
			);
		});

		const addButton = container.querySelector(
			'[data-testid="pptx-master-crud-addLayout"]',
		) as HTMLButtonElement;
		act(() => addButton.click());
		expect(onCrudAction).toHaveBeenCalledWith('addLayout');
	});
});
