import express, { type Request, type Response } from 'express';
import { SSEServerTransport } from '@modelcontextprotocol/sdk/server/sse.js';
import { GmailMultiInboxServer } from './app.js';


function resolveTransportMode(): 'stdio' | 'sse' {
  const explicit = process.env.MCP_TRANSPORT?.trim().toLowerCase();
  if (explicit === 'stdio' || explicit === 'sse') return explicit;
  if (process.env.PORT?.trim()) return 'sse';
  return 'stdio';
}

async function runSseServer(): Promise<void> {
  const port = Number.parseInt(process.env.PORT ?? process.env.RAILWAY_PORT ?? '3000', 10);
  if (!Number.isFinite(port) || port <= 0) {
    throw new Error('Invalid PORT value: ' + (process.env.PORT ?? '(missing)'));
  }

  const host = process.env.HOST?.trim() || '0.0.0.0';
  const app = express();
  const sessions = new Map<string, { app: GmailMultiInboxServer; transport: SSEServerTransport }>();

  const closeAll = async (): Promise<void> => {
    const entries = [...sessions.values()];
    sessions.clear();
    await Promise.allSettled(entries.map(async ({ app, transport }) => {
      await Promise.allSettled([transport.close(), app.close()]);
    }));
  };

  app.get('/', (_req: Request, res: Response) => {
    res.status(200).type('text/plain').send('ghub SSE server');
  });

  app.use((req: Request, _res: Response, next) => {
    console.error('[ghub] HTTP ' + req.method + ' ' + req.path);
    next();
  });

  app.get('/sse', async (_req: Request, res: Response) => {
    const serverApp = new GmailMultiInboxServer();
    const transport = new SSEServerTransport('/messages', res);
    const sessionId = transport.sessionId;
    sessions.set(sessionId, { app: serverApp, transport });

    transport.onclose = () => {
      sessions.delete(sessionId);
    };

    try {
      await serverApp.connectTransport(transport);
      console.error('[ghub] SSE session started: ' + sessionId);
    } catch (error) {
      sessions.delete(sessionId);
      if (!res.headersSent) {
        res.status(500).type('text/plain');
      }
      res.end(error instanceof Error ? error.message : String(error));
    }
  });

  app.post('/messages', async (req: Request, res: Response) => {
    const sessionId = typeof req.query.sessionId === 'string' ? req.query.sessionId : '';
    const session = sessions.get(sessionId);
    if (!session) {
      res.status(404).type('text/plain').send('Unknown SSE session');
      return;
    }

    await session.transport.handlePostMessage(req, res);
  });

  const httpServer = app.listen(port, () => {
    console.error('[ghub] Running on SSE at port ' + port + '. Routes: GET /sse, POST /messages');
  });

  const shutdown = async (): Promise<void> => {
    await closeAll();
    await new Promise<void>((resolve) => httpServer.close(() => resolve()));
  };

  process.on('SIGINT', async () => {
    await shutdown();
    process.exit(0);
  });
  process.on('SIGTERM', async () => {
    await shutdown();
    process.exit(0);
  });
}

export async function runMcp(): Promise<void> {
  if (resolveTransportMode() === 'sse') {
    await runSseServer();
    return;
  }

  const server = new GmailMultiInboxServer();
  process.on('SIGINT', async () => {
    await server.close();
    process.exit(0);
  });
  process.on('SIGTERM', async () => {
    await server.close();
    process.exit(0);
  });
  await server.run();
}
