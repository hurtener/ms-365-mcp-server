import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import GraphClient from '../src/graph-client.js';
import { requestContext } from '../src/request-context.js';
import { microsoftBearerTokenAuthMiddleware } from '../src/lib/microsoft-auth.js';
import { MicrosoftOAuthProvider } from '../src/oauth-provider.js';

vi.mock('../src/logger.js', () => ({
  default: {
    info: vi.fn(),
    error: vi.fn(),
    warn: vi.fn(),
  },
}));

describe('HTTP stateless auth', () => {
  let originalFetch: typeof global.fetch;

  beforeEach(() => {
    originalFetch = global.fetch;
  });

  afterEach(() => {
    global.fetch = originalFetch;
    vi.restoreAllMocks();
  });

  it('requires a request-scoped token in HTTP mode and does not fall back to cached auth', async () => {
    const authManager = {
      getToken: vi.fn().mockResolvedValue('GLOBAL_DEVICE_TOKEN'),
    };
    const graphClient = new GraphClient(
      authManager as never,
      {
        clientId: 'client-id',
        tenantId: 'common',
        cloudType: 'global',
      },
      'json',
      true
    );

    await expect(graphClient.makeRequest('/me')).rejects.toThrow(
      'Missing request-scoped access token for HTTP mode'
    );
    expect(authManager.getToken).not.toHaveBeenCalled();
  });

  it('uses the request-scoped bearer token when present in HTTP mode', async () => {
    let capturedToken: string | undefined;
    global.fetch = vi
      .fn()
      .mockImplementation(async (_url: string, options: { headers?: Record<string, string> }) => {
        capturedToken = options.headers?.Authorization?.replace('Bearer ', '');
        return {
          ok: true,
          status: 200,
          text: async () => JSON.stringify({ id: 'ok' }),
          headers: new Headers(),
        };
      });

    const authManager = {
      getToken: vi.fn().mockResolvedValue('GLOBAL_DEVICE_TOKEN'),
    };
    const graphClient = new GraphClient(
      authManager as never,
      {
        clientId: 'client-id',
        tenantId: 'common',
        cloudType: 'global',
      },
      'json',
      true
    );

    await requestContext.run({ accessToken: 'REQUEST_TOKEN' }, async () => {
      await graphClient.makeRequest('/me');
    });

    expect(capturedToken).toBe('REQUEST_TOKEN');
    expect(authManager.getToken).not.toHaveBeenCalled();
  });

  it('rejects missing Authorization header in HTTP middleware', () => {
    const req = { headers: {} } as never;
    const res = {
      status: vi.fn().mockReturnThis(),
      json: vi.fn(),
    } as never;
    const next = vi.fn();

    microsoftBearerTokenAuthMiddleware(req, res, next);

    expect(res.status).toHaveBeenCalledWith(401);
    expect(res.json).toHaveBeenCalledWith({ error: 'Missing or invalid access token' });
    expect(next).not.toHaveBeenCalled();
  });

  it('verifies bearer tokens without mutating shared auth manager state', async () => {
    global.fetch = vi.fn().mockResolvedValue({
      ok: true,
      json: async () => ({ userPrincipalName: 'juan@company.com' }),
    });

    const provider = new MicrosoftOAuthProvider({
      clientId: 'client-id',
      tenantId: 'common',
      cloudType: 'global',
    });

    const authInfo = await provider.verifyAccessToken('REQUEST_TOKEN');

    expect(authInfo.token).toBe('REQUEST_TOKEN');
  });
});
