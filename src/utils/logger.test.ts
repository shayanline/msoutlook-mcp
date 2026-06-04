import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';

describe('logger', () => {
  let writeSpy: ReturnType<typeof vi.spyOn>;

  beforeEach(() => {
    vi.resetModules();
    writeSpy = vi.spyOn(process.stderr, 'write').mockImplementation(() => true);
  });

  afterEach(() => {
    vi.unstubAllEnvs();
    vi.restoreAllMocks();
  });

  async function load() {
    return (await import('./logger.js')).logger;
  }

  it('info writes with prefix and no args suffix', async () => {
    const logger = await load();
    logger.info('hello');
    expect(writeSpy).toHaveBeenCalledWith('[msoutlook-mcp] hello\n');
  });

  it('info appends JSON when args are present', async () => {
    const logger = await load();
    logger.info('hi', 1, { a: 2 });
    expect(writeSpy).toHaveBeenCalledWith(`[msoutlook-mcp] hi ${JSON.stringify([1, { a: 2 }])}\n`);
  });

  it('warn writes with the warn prefix', async () => {
    const logger = await load();
    logger.warn('careful');
    expect(writeSpy).toHaveBeenCalledWith('[msoutlook-mcp:warn] careful\n');
  });

  it('error writes with the error prefix and args', async () => {
    const logger = await load();
    logger.error('boom', 'x');
    expect(writeSpy).toHaveBeenCalledWith(`[msoutlook-mcp:error] boom ${JSON.stringify(['x'])}\n`);
  });

  it('warn and error omit the JSON suffix when called with no args', async () => {
    const logger = await load();
    logger.warn('w');
    logger.error('e');
    expect(writeSpy).toHaveBeenCalledWith('[msoutlook-mcp:warn] w\n');
    expect(writeSpy).toHaveBeenCalledWith('[msoutlook-mcp:error] e\n');
  });

  it('debug stays silent when MSOUTLOOK_DEBUG is not "true"', async () => {
    vi.stubEnv('MSOUTLOOK_DEBUG', 'false');
    const logger = await load();
    logger.debug('quiet');
    expect(writeSpy).not.toHaveBeenCalled();
  });

  it('debug writes when MSOUTLOOK_DEBUG is "true"', async () => {
    vi.stubEnv('MSOUTLOOK_DEBUG', 'true');
    const logger = await load();
    logger.debug('loud', 1);
    expect(writeSpy).toHaveBeenCalledWith(`[msoutlook-mcp:debug] loud ${JSON.stringify([1])}\n`);
  });

  it('debug with no args and debug enabled omits the JSON suffix', async () => {
    vi.stubEnv('MSOUTLOOK_DEBUG', 'true');
    const logger = await load();
    logger.debug('plain');
    expect(writeSpy).toHaveBeenCalledWith('[msoutlook-mcp:debug] plain\n');
  });
});
