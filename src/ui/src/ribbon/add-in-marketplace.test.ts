import { expect, it } from 'vitest';
import { marketplaceAssetId } from './add-in-marketplace';

it.each([
	' wa999999999 ',
	'https://appsource.microsoft.com/en-us/product/office/WA999999999?tab=overview',
	'https://marketplace.microsoft.com/en-us/product/office/WA999999999',
	'https://store.office.com/en-001/app.aspx?assetid=WA999999999',
	'https://go.microsoft.com/fwlink/?linkid=2261819&templateid=WA999999999',
])('reduces %s to only the asset ID', (value) => {
	expect(marketplaceAssetId(value)).toBe('WA999999999');
});

it.each([
	'http://appsource.microsoft.com/en-us/product/office/WA999999999',
	'https://appsource.microsoft.com.example.test/en-us/product/office/WA999999999',
	'https://user@appsource.microsoft.com/en-us/product/office/WA999999999',
	'https://store.office.com/app.aspx?assetid=WA1&assetid=WA2',
	'https://appsource.microsoft.com/en-us/product/teams/WA999999999',
])('refuses %s before calling a resolver', (value) => {
	expect(marketplaceAssetId(value)).toBeUndefined();
});
