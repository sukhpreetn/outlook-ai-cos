export type LogLevel = 'debug' | 'info' | 'warn' | 'error';
export declare const Logger: {
    debug: (msg: string, data?: unknown) => void;
    info: (msg: string, data?: unknown) => void;
    warn: (msg: string, data?: unknown) => void;
    error: (msg: string, data?: unknown) => void;
    diagIncrement(key: string, by?: number): void;
    diagGet(key: string): number;
    diagReset(key: string): void;
    diagSnapshot(): Record<string, number>;
};
//# sourceMappingURL=Logger.d.ts.map