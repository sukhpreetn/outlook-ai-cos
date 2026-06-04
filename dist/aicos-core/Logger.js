"use strict";
// =============================================================================
// src/aicos-core/Logger.ts
//
// Centralized logging. All modules use Logger, never console.log directly.
// In Azure Functions this writes to Application Insights automatically
// when APPLICATIONINSIGHTS_CONNECTION_STRING is set.
// =============================================================================
Object.defineProperty(exports, "__esModule", { value: true });
exports.Logger = void 0;
// ── Counters for diagnostic reporting (mirrors DIAG_* in original AICOS) ──────
const _counters = {};
exports.Logger = {
    debug: (msg, data) => _log('debug', msg, data),
    info: (msg, data) => _log('info', msg, data),
    warn: (msg, data) => _log('warn', msg, data),
    error: (msg, data) => _log('error', msg, data),
    // Increment a named diagnostic counter (e.g. 'DIAG_AI_ERRORS')
    diagIncrement(key, by = 1) {
        _counters[key] = (_counters[key] ?? 0) + by;
    },
    diagGet(key) {
        return _counters[key] ?? 0;
    },
    diagReset(key) {
        _counters[key] = 0;
    },
    // Dump all counters to a plain object (for daily diagnostic email)
    diagSnapshot() {
        return { ..._counters };
    },
};
function _log(level, msg, data) {
    const entry = {
        ts: new Date().toISOString(),
        level,
        msg,
        ...(data !== undefined ? { data } : {}),
    };
    const line = `[${entry.ts}] [${level.toUpperCase()}] ${msg}`;
    switch (level) {
        case 'debug':
            console.debug(line, data ?? '');
            break;
        case 'info':
            console.info(line, data ?? '');
            break;
        case 'warn':
            console.warn(line, data ?? '');
            break;
        case 'error':
            console.error(line, data ?? '');
            break;
    }
}
//# sourceMappingURL=Logger.js.map