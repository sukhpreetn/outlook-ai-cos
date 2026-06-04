// =============================================================================
// src/aicos-core/Logger.ts
//
// Centralized logging. All modules use Logger, never console.log directly.
// In Azure Functions this writes to Application Insights automatically
// when APPLICATIONINSIGHTS_CONNECTION_STRING is set.
// =============================================================================

export type LogLevel = 'debug' | 'info' | 'warn' | 'error';

interface LogEntry {
  ts:      string;
  level:   LogLevel;
  msg:     string;
  data?:   unknown;
}

// ── Counters for diagnostic reporting (mirrors DIAG_* in original AICOS) ──────
const _counters: Record<string, number> = {};

export const Logger = {
  debug: (msg: string, data?: unknown) => _log('debug', msg, data),
  info:  (msg: string, data?: unknown) => _log('info',  msg, data),
  warn:  (msg: string, data?: unknown) => _log('warn',  msg, data),
  error: (msg: string, data?: unknown) => _log('error', msg, data),

  // Increment a named diagnostic counter (e.g. 'DIAG_AI_ERRORS')
  diagIncrement(key: string, by = 1): void {
    _counters[key] = (_counters[key] ?? 0) + by;
  },

  diagGet(key: string): number {
    return _counters[key] ?? 0;
  },

  diagReset(key: string): void {
    _counters[key] = 0;
  },

  // Dump all counters to a plain object (for daily diagnostic email)
  diagSnapshot(): Record<string, number> {
    return { ..._counters };
  },
};

function _log(level: LogLevel, msg: string, data?: unknown): void {
  const entry: LogEntry = {
    ts:    new Date().toISOString(),
    level,
    msg,
    ...(data !== undefined ? { data } : {}),
  };

  const line = `[${entry.ts}] [${level.toUpperCase()}] ${msg}`;

  switch (level) {
    case 'debug': console.debug(line, data ?? ''); break;
    case 'info':  console.info(line,  data ?? ''); break;
    case 'warn':  console.warn(line,  data ?? ''); break;
    case 'error': console.error(line, data ?? ''); break;
  }
}
