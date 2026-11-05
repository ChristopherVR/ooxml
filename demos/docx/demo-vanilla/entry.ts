import { completeDemoSignInPopup } from 'ooxml-demo-onedrive';

// The MSAL sign-in popup loads this same page: hand the response back and do not boot the demo.
if (!completeDemoSignInPopup()) void import('./src');
