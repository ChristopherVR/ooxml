// @vitest-environment jsdom
import { expect, it } from 'vitest';
import { addInLauncher } from './add-in-launch';
import { readLocalAddIn } from './add-in-manifest-file';

it('refuses to open an approved origin without a host frame policy', async () => {
	const host = document.body.appendChild(document.createElement('div'));
	try {
		const source = `<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
			xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
			<Id>083a367e-842b-4d4f-ab91-cd5de849e495</Id><Version>1.0.0</Version>
			<ProviderName>Notes</ProviderName><DefaultLocale>en</DefaultLocale>
			<DisplayName DefaultValue="Notes"/><Description DefaultValue="Notes"/>
			<DefaultSettings><SourceLocation DefaultValue="https://publisher.example.test/pane"/></DefaultSettings>
			<Permissions>ReadDocument</Permissions>
		</OfficeApp>`;
		const addIn = await readLocalAddIn(new TextEncoder().encode(source), { host: 'word' });
		Reflect.set(host, Symbol.for('office-ui.add-in-runtime'), {
			origins: ['https://publisher.example.test'],
		});
		const launcher = addInLauncher(
			{
				host,
				mode: 'myAddIns',
				t: (text) => text,
				product: 'word',
				manifestLocale: 'en',
				scopes: [host],
			},
			new Map([[addIn.manifest.id, addIn]]),
			(text) => text,
		);
		expect(launcher.pane(addIn).reason).toContain('enforce a frame policy');
		await launcher.open(addIn);
		expect(host.querySelector('office-ui-add-in-pane')).toBeNull();
	} finally {
		host.remove();
	}
});
