import type {
  LLMProvider,
  LLMMessage,
  ToolDefinition,
  GenerateOptions,
  GenerateResult,
} from '../types/index.js';

export interface ProviderConfig {
  apiKey: string;
  baseUrl: string;
  model: string;
}

export abstract class BaseLLMProvider implements LLMProvider {
  abstract name: string;
  protected config: ProviderConfig;

  constructor(config: ProviderConfig) {
    this.config = config;
  }

  abstract generateText(
    messages: LLMMessage[],
    options?: GenerateOptions
  ): Promise<GenerateResult>;

  abstract generateToolCall(
    messages: LLMMessage[],
    tools: ToolDefinition[],
    options?: GenerateOptions
  ): Promise<GenerateResult>;
}
