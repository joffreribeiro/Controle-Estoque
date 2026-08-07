import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';

// debounce.js não usa `export` (script solto, carregado via <script> no index.html) —
// ele só expõe window.Debounce. Como o ambiente de teste é Node (sem jsdom, ver
// vitest.config.js), simulamos `window` antes de importar, igual ao padrão que o
// próprio arquivo já prevê (`typeof window !== 'undefined'`).
beforeEach(() => {
  global.window = global;
  vi.useFakeTimers();
});

afterEach(() => {
  vi.useRealTimers();
  delete global.window;
  vi.resetModules();
});

describe('Debounce.criar', () => {
  it('só executa a função uma vez, após o delay, quando chamada várias vezes seguidas', async () => {
    await import('../debounce.js');
    const fn = vi.fn();
    const debounced = window.Debounce.criar(fn, 300);

    debounced();
    debounced();
    debounced();
    expect(fn).not.toHaveBeenCalled();

    vi.advanceTimersByTime(299);
    expect(fn).not.toHaveBeenCalled();

    vi.advanceTimersByTime(1);
    expect(fn).toHaveBeenCalledTimes(1);
  });

  it('reinicia o timer a cada nova chamada (não executa antes do delay contado da última)', async () => {
    await import('../debounce.js');
    const fn = vi.fn();
    const debounced = window.Debounce.criar(fn, 300);

    debounced();
    vi.advanceTimersByTime(200);
    debounced(); // reinicia
    vi.advanceTimersByTime(200);
    expect(fn).not.toHaveBeenCalled();

    vi.advanceTimersByTime(100);
    expect(fn).toHaveBeenCalledTimes(1);
  });

  it('repassa os argumentos da última chamada para a função original', async () => {
    await import('../debounce.js');
    const fn = vi.fn();
    const debounced = window.Debounce.criar(fn, 100);

    debounced('a', 1);
    debounced('b', 2);
    vi.advanceTimersByTime(100);

    expect(fn).toHaveBeenCalledTimes(1);
    expect(fn).toHaveBeenCalledWith('b', 2);
  });

  it('usa 250ms como padrão quando delayMs não é informado', async () => {
    await import('../debounce.js');
    const fn = vi.fn();
    const debounced = window.Debounce.criar(fn);

    debounced();
    vi.advanceTimersByTime(249);
    expect(fn).not.toHaveBeenCalled();
    vi.advanceTimersByTime(1);
    expect(fn).toHaveBeenCalledTimes(1);
  });
});
