import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import tailwindcss from '@tailwindcss/vite';

// base: './' uses relative asset paths in the built output, so the site
// works regardless of which GitHub Pages subpath it ends up deployed under
// (https://<user>.github.io/<any-repo-name>/) without needing to hardcode
// the repo name here.
export default defineConfig({
  base: './',
  plugins: [react(), tailwindcss()]
});
