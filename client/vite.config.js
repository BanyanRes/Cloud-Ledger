import { defineConfig, loadEnv } from 'vite';
import react from '@vitejs/plugin-react';
import { fileURLToPath } from 'url';

// Dev-server ports come from the repo-root .env (PORT = API server,
// CLIENT_PORT = this Vite server) so parallel worktrees don't collide.
// Shell env vars override .env. Defaults match the old hardcoded values.
// Only affects 'vite' dev mode; 'vite build' (production) ignores server.*.
export default defineConfig(({ mode }) => {
  const rootDir = fileURLToPath(new URL('..', import.meta.url));
  const env = { ...loadEnv(mode, rootDir, ''), ...process.env };
  const apiPort = env.PORT || 3000;
  const clientPort = Number(env.CLIENT_PORT) || 5173;
  return {
    plugins: [react()],
    server: {
      port: clientPort,
      strictPort: true,
      proxy: {
        '/api': `http://localhost:${apiPort}`
      }
    }
  };
});
