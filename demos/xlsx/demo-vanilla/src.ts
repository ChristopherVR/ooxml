// The page stylesheet loads with this entry, not with the lazily imported demo below: imported
// only from `app.ts`, the landing page painted as raw HTML until the whole editor had loaded.
import './style.css';
import { completeDemoSignInPopup } from 'ooxml-demo-onedrive';

// In the MSAL sign-in popup the page only hands the response back; the demo must not boot there.
if (!completeDemoSignInPopup()) void import('./app');
