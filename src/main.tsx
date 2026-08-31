import {StrictMode} from 'react';
import {createRoot} from 'react-dom/client';
import App from './App.tsx';
import './index.css';
import { GoogleOAuthProvider } from '@react-oauth/google';

// Note: The OAuth Client ID is public and can be safely hardcoded or provided via Vercel Env Vars.
const clientId = import.meta.env.VITE_GOOGLE_CLIENT_ID || '483256151712-akduq8v312k2vv569cauo9h6rvbhgq69.apps.googleusercontent.com';

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <GoogleOAuthProvider clientId={clientId}>
      <App />
    </GoogleOAuthProvider>
  </StrictMode>,
);
