import express from 'express';
import { createServer as createViteServer } from 'vite';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

async function startServer() {
  const app = express();
  const PORT = Number(process.env.PORT) || 3000;

  // Endpoint to proxy SharePoint / OneDrive / Google Sheet files to bypass browser CORS
  app.get('/api/proxy-excel', async (req, res) => {
    const targetUrl = req.query.url as string;
    if (!targetUrl) {
      return res.status(400).json({ message: 'Parameter url diperlukan.' });
    }

    try {
      const response = await fetch(targetUrl, {
        method: 'GET',
        headers: {
          'User-Agent':
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
          Accept:
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel, text/csv, */*'
        },
        redirect: 'follow'
      });

      if (response.status === 401 || response.status === 403) {
        return res.status(403).json({
          message:
            'File belum di-share ke "Anyone with link". Silahkan ubah permission di SharePoint agar dapat diakses tanpa login.'
        });
      }

      if (!response.ok) {
        return res
          .status(response.status)
          .json({ message: `SharePoint server mengembalikan status ${response.status}.` });
      }

      const buffer = await response.arrayBuffer();

      // Check if response is HTML login page
      const sample = Buffer.from(buffer.slice(0, 3000)).toString('utf-8').toLowerCase();
      if (
        sample.includes('login.microsoftonline.com') ||
        sample.includes('account.live.com') ||
        sample.includes('sign in to your account') ||
        sample.includes('aadsts') ||
        (sample.includes('<!doctype html') && sample.includes('microsoft'))
      ) {
        return res.status(403).json({
          message:
            'File belum di-share ke "Anyone with link". Silahkan ubah permission di SharePoint agar dapat diakses tanpa login.'
        });
      }

      res.setHeader(
        'Content-Type',
        response.headers.get('content-type') ||
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      );
      res.setHeader('Cache-Control', 'no-cache');
      return res.send(Buffer.from(buffer));
    } catch (err: any) {
      console.error('Proxy Error:', err);
      return res.status(500).json({
        message:
          'File belum di-share ke "Anyone with link", silahkan ubah permission di SharePoint.'
      });
    }
  });

  // Health endpoint
  app.get('/api/health', (_req, res) => {
    res.json({ status: 'ok' });
  });

  const isProd = process.env.NODE_ENV === 'production';
  if (!isProd) {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa'
    });
    app.use(vite.middlewares);
  } else {
    app.use(express.static(path.resolve(__dirname, 'dist')));
    app.get('*', (_req, res) => {
      res.sendFile(path.resolve(__dirname, 'dist', 'index.html'));
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server running on http://0.0.0.0:${PORT}`);
  });
}

startServer();
