// The page stylesheet loads with this entry, not with the lazily imported demo below: imported
// only from `src.ts`, the landing page painted as raw HTML until the whole editor had loaded.
import './style.css';
import { completeDemoSignInPopup } from 'ooxml-demo-onedrive';

// The MSAL sign-in popup loads this same page: hand the response back and do not boot the demo.
if (!completeDemoSignInPopup()) void import('./src');
