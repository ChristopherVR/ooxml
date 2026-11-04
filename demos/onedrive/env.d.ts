/// <reference types="vite/client" />

interface ImportMetaEnv {
	/** Application (client) ID of your Entra app registration (single-page application). */
	readonly VITE_ONEDRIVE_CLIENT_ID?: string;
	/** `consumers` (personal accounts, the default) or `common` (work, school or personal). */
	readonly VITE_ONEDRIVE_AUTHORITY?: string;
}
