export type AITask = 'classify' | 'brief' | 'draft' | 'learn' | 'summarize' | 'route';
export interface AIOptions {
    maxTokens?: number;
    temperature?: number;
    systemPrompt?: string;
}
export interface AIResult {
    text: string;
    model: string;
    task: AITask;
    tokens?: {
        input: number;
        output: number;
    };
}
export declare function callAI(task: AITask, prompt: string, options?: AIOptions): Promise<AIResult>;
//# sourceMappingURL=AI.d.ts.map