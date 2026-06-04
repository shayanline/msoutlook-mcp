import { describe, it, expect, vi, beforeEach } from 'vitest';
import {
  browserLogin, getAuthStatus, clearSession, getOwaToken,
  hasSessionState, isSessionLikelyExpired,
} from '../auth/index.js';
import { registerAuthTools } from './auth-tools.js';

vi.mock('../auth/index.js', () => ({
  browserLogin: vi.fn(),
  getAuthStatus: vi.fn(),
  clearSession: vi.fn(),
  getOwaToken: vi.fn(),
  hasSessionState: vi.fn(),
  isSessionLikelyExpired: vi.fn(),
}));

type Reg = { schema: any; handler: (a: any) => Promise<any> };
let tools: Map<string, Reg>;
function setup() {
  tools = new Map();
  const server = { tool: (n: string, _d: string, s: any, h: any) => tools.set(n, { schema: s, handler: h }) } as any;
  registerAuthTools(server);
}
const text = (r: any) => r.content[0].text;

beforeEach(() => { vi.clearAllMocks(); setup(); });

describe('outlook_login', () => {
  it('fast path: valid token already authenticated', async () => {
    vi.mocked(getOwaToken).mockResolvedValue('tok');
    vi.mocked(getAuthStatus).mockReturnValue({ upn: 'me@x.com', owaTokenMinutesRemaining: 30, owaTokenExpiry: 'e' } as any);
    const p = JSON.parse(text(await tools.get('outlook_login')!.handler({})));
    expect(p.success).toBe(true);
    expect(p.message).toContain('30 more minutes');
    expect(browserLogin).not.toHaveBeenCalled();
  });
  it('token without minutes defaults 0 then browser login', async () => {
    vi.mocked(getOwaToken).mockResolvedValue('tok');
    vi.mocked(getAuthStatus).mockReturnValue({ upn: 'me@x.com' } as any);
    vi.mocked(browserLogin).mockResolvedValue({ method: 'headed-browser', upn: 'me@x.com' } as any);
    const p = JSON.parse(text(await tools.get('outlook_login')!.handler({})));
    expect(p.success).toBe(true);
    expect(browserLogin).toHaveBeenCalledWith(false);
    expect(p.message).toContain('successfully');
  });
  it('token below threshold -> headless-sso', async () => {
    vi.mocked(getOwaToken).mockResolvedValue('tok');
    vi.mocked(getAuthStatus).mockReturnValue({ owaTokenMinutesRemaining: 5 } as any);
    vi.mocked(browserLogin).mockResolvedValue({ method: 'headless-sso', upn: 'a@b.com' } as any);
    const p = JSON.parse(text(await tools.get('outlook_login')!.handler({})));
    expect(p.message).toContain('silently');
  });
  it('no token -> token-cache message, writes stderr', async () => {
    vi.mocked(getOwaToken).mockResolvedValue(null);
    vi.mocked(browserLogin).mockResolvedValue({ method: 'token-cache', upn: 'c@d.com' } as any);
    const spy = vi.spyOn(process.stderr, 'write').mockReturnValue(true);
    const p = JSON.parse(text(await tools.get('outlook_login')!.handler({})));
    expect(p.message).toContain('Already authenticated');
    expect(spy).toHaveBeenCalled();
    spy.mockRestore();
  });
  it('force_new skips fast path -> browserLogin(true)', async () => {
    vi.mocked(browserLogin).mockResolvedValue({ method: 'headed-browser', upn: 'f@g.com' } as any);
    const p = JSON.parse(text(await tools.get('outlook_login')!.handler({ force_new: true })));
    expect(getOwaToken).not.toHaveBeenCalled();
    expect(browserLogin).toHaveBeenCalledWith(true);
    expect(p.success).toBe(true);
  });
  it('login failure -> success false', async () => {
    vi.mocked(getOwaToken).mockResolvedValue(null);
    vi.mocked(browserLogin).mockResolvedValue(null as any);
    vi.spyOn(process.stderr, 'write').mockReturnValue(true);
    const p = JSON.parse(text(await tools.get('outlook_login')!.handler({})));
    expect(p.success).toBe(false);
    expect(p.message).toContain('Login failed');
  });
});

describe('outlook_status', () => {
  it('not authenticated', async () => {
    vi.mocked(getOwaToken).mockResolvedValue(null);
    const p = JSON.parse(text(await tools.get('outlook_status')!.handler({})));
    expect(p.authenticated).toBe(false);
  });
  it('full status with graph token', async () => {
    vi.mocked(getOwaToken).mockResolvedValue('tok');
    vi.mocked(getAuthStatus).mockReturnValue({ upn: 'm', tenantId: 't', owaTokenExpiry: 'e', owaTokenMinutesRemaining: 50, graphTokenExpiry: 'g' } as any);
    vi.mocked(hasSessionState).mockReturnValue(true);
    vi.mocked(isSessionLikelyExpired).mockReturnValue(false);
    const p = JSON.parse(text(await tools.get('outlook_status')!.handler({})));
    expect(p.authenticated).toBe(true);
    expect(p.graphToken).toEqual({ expiresAt: 'g' });
    expect(p.session).toEqual({ exists: true, likelyExpired: false });
  });
  it('null graph token', async () => {
    vi.mocked(getOwaToken).mockResolvedValue('tok');
    vi.mocked(getAuthStatus).mockReturnValue({ upn: 'm' } as any);
    vi.mocked(hasSessionState).mockReturnValue(false);
    vi.mocked(isSessionLikelyExpired).mockReturnValue(true);
    const p = JSON.parse(text(await tools.get('outlook_status')!.handler({})));
    expect(p.graphToken).toBeNull();
  });
});

describe('outlook_logout', () => {
  it('clears session', async () => {
    const p = JSON.parse(text(await tools.get('outlook_logout')!.handler({})));
    expect(clearSession).toHaveBeenCalled();
    expect(p.success).toBe(true);
  });
});
