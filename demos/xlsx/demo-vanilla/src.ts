import { completeDemoSignInPopup } from 'ooxml-demo-onedrive';

// In the MSAL sign-in popup the page only hands the response back; the demo must not boot there.
if (!completeDemoSignInPopup()) void import('./app');
