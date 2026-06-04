import { getConfig } from './Config.js';
import { Logger } from './Logger.js';

export type AITask = 'classify' | 'brief' | 'draft' | 'learn' | 'summarize' | 'route';

export interface AIOptions {
  maxTokens?:   number;
  temperature?: number;
  systemPrompt?: string;
}

export interface AIResult {
  text:    string;
  model:   string;
  task:    AITask;
  tokens?: { input: number; output: number };
}

const TASK_DEFAULTS: Record<AITask, string> = {
  classify:  'claude-haiku-4-5',
  learn:     'claude-haiku-4-5',
  draft:     'claude-haiku-4-5',
  route:     'claude-haiku-4-5',
  brief:     'claude-sonnet-4-6',
  summarize: 'claude-sonnet-4-6',
};

export async function callAI(
  task: AITask,
  prompt: string,
  options: AIOptions = {},
): Promise<AIResult> {
  const model = TASK_DEFAULTS[task];
  Logger.info(`[AI] task=${task} model=${model}`);

  const body = {
    model,
    max_tokens:  options.maxTokens  ?? _defaultMaxTokens(task),
    temperature: options.temperature ?? _defaultTemperature(task),
    system:      options.systemPrompt ?? _defaultSystemPrompt(task),
    messages: [{ role: 'user', content: prompt }],
  };

  const cfg = await getConfig();
  const resp = await fetch('https://api.anthropic.com/v1/messages', {
    method:  'POST',
    headers: {
      'Content-Type':      'application/json',
      'x-api-key':         cfg.anthropicApiKey,
      'anthropic-version': '2023-06-01',
    },
    body: JSON.stringify(body),
  });

  if (!resp.ok) {
    const err = await resp.text();
    throw new Error(`Anthropic API error ${resp.status}: ${err}`);
  }

  const data = await resp.json() as {
    content: Array<{ type: string; text?: string }>;
    usage: { input_tokens: number; output_tokens: number };
  };

  const text = data.content
    .filter(b => b.type === 'text')
    .map(b => b.text ?? '')
    .join('');

  Logger.info(`[AI] done task=${task} in=${data.usage.input_tokens} out=${data.usage.output_tokens}`);

  return { text, model, task, tokens: { input: data.usage.input_tokens, output: data.usage.output_tokens } };
}

function _defaultMaxTokens(task: AITask): number {
  const map: Record<AITask, number> = { classify: 512, learn: 512, draft: 1024, route: 256, brief: 4096, summarize: 2048 };
  return map[task];
}

function _defaultTemperature(task: AITask): number {
  const map: Record<AITask, number> = { classify: 0.0, learn: 0.0, route: 0.0, draft: 0.7, brief: 0.5, summarize: 0.3 };
  return map[task];
}

function _defaultSystemPrompt(task: AITask): string {
  const base = "You are an AI Chief of Staff managing an executive's Outlook inbox.";
  const map: Record<AITask, string> = {
    classify:  `${base} Classify emails accurately. Respond only with valid JSON.`,
    learn:     `${base} Analyze patterns and extract rules. Respond only with valid JSON.`,
    route:     `${base} Determine routing. Respond only with valid JSON.`,
    draft:     `${base} Draft concise, professional replies matching the user's tone.`,
    brief:     `${base} Write a crisp, actionable daily briefing. Be direct and valuable.`,
    summarize: `${base} Summarize thoroughly and extract key insights and required actions.`,
  };
  return map[task];
}
