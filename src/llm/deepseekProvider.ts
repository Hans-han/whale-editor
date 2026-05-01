import { BaseLLMProvider, type ProviderConfig } from './provider.js';
import { extractUsage } from './usage.js';
import { logger } from '../utils/logger.js';
import type {
  LLMMessage,
  ToolDefinition,
  GenerateOptions,
  GenerateResult,
  ToolCall,
} from '../types/index.js';

interface ChatMessage {
  role: string;
  content: string | null;
  tool_calls?: Array<{
    id: string;
    type: 'function';
    function: { name: string; arguments: string };
  }>;
  tool_call_id?: string;
}

interface ChatChoice {
  message: ChatMessage;
  finish_reason: string;
}

interface ChatResponse {
  id: string;
  choices: ChatChoice[];
  usage: {
    prompt_tokens: number;
    completion_tokens: number;
    total_tokens: number;
    prompt_cache_hit_tokens?: number;
    prompt_cache_miss_tokens?: number;
  };
}

export interface DeepSeekProviderOverride {
  apiKey?: string;
  baseUrl?: string;
  model?: string;
}

export class DeepSeekProvider extends BaseLLMProvider {
  name = 'deepseek';

  constructor(override?: DeepSeekProviderOverride) {
    const apiKey = override?.apiKey ?? process.env.DEEPSEEK_API_KEY ?? '';
    const baseUrl =
      override?.baseUrl ?? process.env.DEEPSEEK_BASE_URL ?? 'https://api.deepseek.com/v1';
    const model = override?.model ?? process.env.DEEPSEEK_MODEL ?? 'deepseek-v4';

    if (!apiKey) {
      logger.warn('DEEPSEEK_API_KEY not set. API calls will fail.');
    }

    super({ apiKey, baseUrl, model });
  }

  async generateText(
    messages: LLMMessage[],
    options?: GenerateOptions
  ): Promise<GenerateResult> {
    return this.chatCompletion(messages, undefined, options);
  }

  async generateToolCall(
    messages: LLMMessage[],
    tools: ToolDefinition[],
    options?: GenerateOptions
  ): Promise<GenerateResult> {
    return this.chatCompletion(messages, tools, {
      ...options,
      toolChoice: options?.toolChoice ?? 'auto',
    });
  }

  private async chatCompletion(
    messages: LLMMessage[],
    tools?: ToolDefinition[],
    options?: GenerateOptions
  ): Promise<GenerateResult> {
    const body: Record<string, unknown> = {
      model: this.config.model,
      messages: messages.map(this.formatMessage),
      temperature: options?.temperature ?? 0.1,
      max_tokens: options?.maxTokens ?? 4096,
    };

    if (tools && tools.length > 0) {
      body.tools = tools.map((t) => ({
        type: 'function',
        function: {
          name: t.function.name,
          description: t.function.description,
          parameters: t.function.parameters,
        },
      }));

      if (options?.toolChoice) {
        body.tool_choice = options.toolChoice;
      }
    }

    const url = `${this.config.baseUrl}/chat/completions`;

    logger.debug(`Calling ${url} with model ${this.config.model}`);

    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${this.config.apiKey}`,
      },
      body: JSON.stringify(body),
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`DeepSeek API error ${response.status}: ${errorText}`);
    }

    const data = (await response.json()) as ChatResponse;
    const choice = data.choices?.[0];

    if (!choice) {
      throw new Error('No choice returned from DeepSeek API');
    }

    const usage = extractUsage(data.usage);
    logger.info('LLM usage', {
      inputTokens: usage.inputTokens,
      outputTokens: usage.outputTokens,
      cacheHitTokens: usage.cacheHitTokens,
      cacheMissTokens: usage.cacheMissTokens,
      cacheHitRatio: usage.cacheHitRatio,
    });

    const toolCalls = this.extractToolCalls(choice.message);

    return {
      content: choice.message.content ?? null,
      toolCalls: toolCalls.length > 0 ? toolCalls : undefined,
      usage,
    };
  }

  private formatMessage(msg: LLMMessage): ChatMessage {
    const formatted: ChatMessage = {
      role: msg.role,
      content: msg.content,
    };

    if (msg.toolCallId) {
      formatted.tool_call_id = msg.toolCallId;
    }

    if (msg.toolCalls) {
      formatted.tool_calls = msg.toolCalls.map((tc) => ({
        id: tc.id,
        type: 'function' as const,
        function: {
          name: tc.function.name,
          arguments: tc.function.arguments,
        },
      }));
    }

    return formatted;
  }

  private extractToolCalls(message: ChatMessage): ToolCall[] {
    if (!message.tool_calls) return [];

    return message.tool_calls.map((tc) => ({
      id: tc.id,
      type: 'function' as const,
      function: {
        name: tc.function.name,
        arguments: tc.function.arguments,
      },
    }));
  }
}
