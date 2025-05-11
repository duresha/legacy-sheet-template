import { defineConfig } from 'vite';
import { resolve } from 'path';
import { viteStaticCopy } from 'vite-plugin-static-copy';
import path from 'path';
import fs from 'fs';

export default defineConfig({
  root: './src',
  publicDir: '../public',
  build: {
    outDir: '../dist',
    rollupOptions: {
      input: {
        main: resolve(__dirname, 'src/index.html'),
      },
    },
  },
  plugins: [
    viteStaticCopy({
      targets: [
        {
          src: '../node_modules/pdfjs-dist/build/pdf.worker.mjs',
          dest: '',
          rename: 'pdf.worker.min.mjs'
        }
      ]
    }),
    {
      name: 'configure-server',
      configureServer(server) {
        server.middlewares.use((req, res, next) => {
          if (req.url === '/api/save-state' && req.method === 'POST') {
            let body = '';
            req.on('data', chunk => {
              body += chunk.toString();
            });
            
            req.on('end', () => {
              try {
                const dataDir = path.resolve('./data');
                if (!fs.existsSync(dataDir)) {
                  fs.mkdirSync(dataDir, { recursive: true });
                }
                
                const statePath = path.resolve('./data/appState.json');
                fs.writeFileSync(statePath, body);
                
                res.statusCode = 200;
                res.setHeader('Content-Type', 'application/json');
                res.end(JSON.stringify({ success: true, message: 'State saved successfully' }));
              } catch (error) {
                console.error('Error saving state:', error);
                res.statusCode = 500;
                res.setHeader('Content-Type', 'application/json');
                res.end(JSON.stringify({ success: false, error: error.message }));
              }
            });
          } 
          else if (req.url === '/api/load-state' && req.method === 'GET') {
            try {
              const statePath = path.resolve('./data/appState.json');
              
              if (fs.existsSync(statePath)) {
                const state = fs.readFileSync(statePath, 'utf8');
                res.statusCode = 200;
                res.setHeader('Content-Type', 'application/json');
                res.end(state);
              } else {
                res.statusCode = 404;
                res.setHeader('Content-Type', 'application/json');
                res.end(JSON.stringify({ success: false, error: 'No saved state found' }));
              }
            } catch (error) {
              console.error('Error loading state:', error);
              res.statusCode = 500;
              res.setHeader('Content-Type', 'application/json');
              res.end(JSON.stringify({ success: false, error: error.message }));
            }
          } else {
            next();
          }
        });
      }
    }
  ],
  resolve: {
    alias: {
      '@': path.resolve(__dirname, './src'),
    }
  },
  server: {
    open: true,
    port: 3001,
  },
});
