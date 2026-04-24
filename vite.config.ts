import { defineConfig } from 'vite';
import { getHttpsServerOptions } from 'office-addin-dev-certs';

export default defineConfig(async () => {
  const httpsOptions = await getHttpsServerOptions();

  return {
    server: {
      host: 'localhost',
      port: 3000,
      https: httpsOptions
    },
    preview: {
      host: 'localhost',
      port: 4173,
      https: httpsOptions
    }
  };
});
