import { App, PluginSettingTab, Setting } from 'obsidian';
import type PdfCanvasAiPlugin from './main';
import type { HighlightColor } from './types/annotations';
import { HIGHLIGHT_COLORS, COLOR_HEX } from './types/annotations';

export type AiProvider =
  | 'openai'
  | 'anthropic'
  | 'google-gemini'
  | 'deepseek'
  | 'groq'
  | 'xai'
  | 'mistral'
  | 'openrouter'
  | 'together-ai'
  | 'fireworks-ai'
  | 'cerebras'
  | 'ollama'
  | 'lmstudio'
  | 'local-proxy'
  | 'custom';

export interface ProviderConfig {
  provider: AiProvider;
  baseUrl: string;
  apiKey: string;
  model: string;
}

export type DictionarySource = 'auto' | 'local' | 'api';
export type ExportFormat = 'callout' | 'plain';

export interface PluginSettings {
  enableAi: boolean;
  provider: AiProvider;
  baseUrl: string;
  apiKey: string;
  model: string;
  maxContextChars: number;
  systemPrompt: string;
  proxyAutoStart: boolean;
  colorLabels: Record<HighlightColor, string>;
  // Reader settings
  defaultZoom: number;
  defaultHighlightColor: HighlightColor;
  resumeLastPage: boolean;
  // Export settings
  exportFormat: ExportFormat;
  // Dictionary
  dictionarySource: DictionarySource;
}

export const DEFAULT_COLOR_LABELS: Record<HighlightColor, string> = {
  yellow: 'Important',
  green: 'Key Point',
  blue: 'Definition',
  pink: 'Question',
  red: 'Disagree',
};

const PROVIDER_DEFAULTS: Record<AiProvider, Omit<ProviderConfig, 'provider'>> = {
  openai: {
    baseUrl: 'https://api.openai.com/v1',
    apiKey: '',
    model: 'gpt-4o',
  },
  anthropic: {
    baseUrl: 'https://api.anthropic.com/v1',
    apiKey: '',
    model: 'claude-sonnet-4-20250514',
  },
  'google-gemini': {
    baseUrl: 'https://generativelanguage.googleapis.com/v1beta/openai',
    apiKey: '',
    model: 'gemini-2.5-pro',
  },
  deepseek: {
    baseUrl: 'https://api.deepseek.com/v1',
    apiKey: '',
    model: 'deepseek-chat',
  },
  groq: {
    baseUrl: 'https://api.groq.com/openai/v1',
    apiKey: '',
    model: 'llama-3.3-70b-versatile',
  },
  xai: {
    baseUrl: 'https://api.x.ai/v1',
    apiKey: '',
    model: 'grok-3',
  },
  mistral: {
    baseUrl: 'https://api.mistral.ai/v1',
    apiKey: '',
    model: 'mistral-large-latest',
  },
  openrouter: {
    baseUrl: 'https://openrouter.ai/api/v1',
    apiKey: '',
    model: 'anthropic/claude-sonnet-4',
  },
  'together-ai': {
    baseUrl: 'https://api.together.xyz/v1',
    apiKey: '',
    model: 'meta-llama/Llama-3.3-70B-Instruct-Turbo',
  },
  'fireworks-ai': {
    baseUrl: 'https://api.fireworks.ai/inference/v1',
    apiKey: '',
    model: 'accounts/fireworks/models/llama-v3p3-70b-instruct',
  },
  cerebras: {
    baseUrl: 'https://api.cerebras.ai/v1',
    apiKey: '',
    model: 'llama-3.3-70b',
  },
  ollama: {
    baseUrl: 'http://localhost:11434/v1',
    apiKey: '',
    model: 'llama3',
  },
  lmstudio: {
    baseUrl: 'http://localhost:1234/v1',
    apiKey: '',
    model: '',
  },
  'local-proxy': {
    baseUrl: 'http://localhost:3456/v1',
    apiKey: '',
    model: 'claude-opus-4',
  },
  custom: {
    baseUrl: '',
    apiKey: '',
    model: '',
  },
};

export const DEFAULT_SETTINGS: PluginSettings = {
  enableAi: true,
  provider: 'openai',
  baseUrl: 'https://api.openai.com/v1',
  apiKey: '',
  model: 'gpt-4o',
  maxContextChars: 80000,
  proxyAutoStart: false,
  systemPrompt:
    'You are a helpful assistant that analyzes PDF documents. ' +
    'When PDF content is provided as context, analyze it carefully and answer questions about it. ' +
    'Be concise and precise. When quoting from the document, use block quotes.',
  colorLabels: { ...DEFAULT_COLOR_LABELS },
  defaultZoom: 1.5,
  defaultHighlightColor: 'yellow',
  resumeLastPage: true,
  exportFormat: 'callout',
  dictionarySource: 'auto',
};

const PROVIDER_LABELS: Record<AiProvider, string> = {
  openai: 'OpenAI',
  anthropic: 'Anthropic (Claude)',
  'google-gemini': 'Google Gemini',
  deepseek: 'DeepSeek',
  groq: 'Groq',
  xai: 'xAI (Grok)',
  mistral: 'Mistral',
  openrouter: 'OpenRouter',
  'together-ai': 'Together AI',
  'fireworks-ai': 'Fireworks AI',
  cerebras: 'Cerebras',
  ollama: 'Ollama (local)',
  lmstudio: 'LM Studio (local)',
  'local-proxy': 'Local Proxy (claude-max-api-proxy)',
  custom: 'Custom OpenAI-compatible',
};

export class PdfCanvasAiSettingTab extends PluginSettingTab {
  plugin: PdfCanvasAiPlugin;

  constructor(app: App, plugin: PdfCanvasAiPlugin) {
    super(app, plugin);
    this.plugin = plugin;
  }

  display(): void {
    const { containerEl } = this;
    containerEl.empty();

    new Setting(containerEl).setName('Basics').setHeading();

    // ── AI toggle ──
    new Setting(containerEl)
      .setName('Enable AI features')
      .setDesc(
        'Turn off to use PDF Tools as a pure reader/annotator without any AI. ' +
        'Hides the AI sidebar, chat commands, "Ask AI" buttons, and annotation summary. ' +
        'Requires reloading the plugin to take full effect.',
      )
      .addToggle((toggle) =>
        toggle
          .setValue(this.plugin.settings.enableAi)
          .onChange(async (value) => {
            this.plugin.settings.enableAi = value;
            await this.plugin.saveSettings();
            this.display(); // re-render to show/hide AI settings
          }),
      );

    // ── AI settings (only shown when AI is enabled) ──
    if (this.plugin.settings.enableAi) {
      new Setting(containerEl).setName('AI configuration').setHeading();

      // ── Provider selector ──
      new Setting(containerEl)
        .setName('AI provider')
        .setDesc('Select how to connect to an AI model.')
        .addDropdown((dd) => {
          for (const [key, label] of Object.entries(PROVIDER_LABELS)) {
            dd.addOption(key, label);
          }
          dd.setValue(this.plugin.settings.provider);
          dd.onChange(async (value) => {
            const provider = value as AiProvider;
            const prevProvider = this.plugin.settings.provider;
            this.plugin.settings.provider = provider;
            const defaults = PROVIDER_DEFAULTS[provider];
            if (!this.plugin.settings.baseUrl || this.plugin.settings.baseUrl === PROVIDER_DEFAULTS[prevProvider].baseUrl) {
              this.plugin.settings.baseUrl = defaults.baseUrl;
            }
            if (!this.plugin.settings.model || this.plugin.settings.model === PROVIDER_DEFAULTS[prevProvider].model) {
              this.plugin.settings.model = defaults.model;
            }
            await this.plugin.saveSettings();
            this.display();
          });
        });

      // ── API key ──
      const provider = this.plugin.settings.provider;
      const isLocal = provider === 'local-proxy' || provider === 'ollama' || provider === 'lmstudio';
      const keyDesc = isLocal
        ? 'Leave empty — local providers do not require an API key.'
        : `API key for ${PROVIDER_LABELS[provider]}.`;
      const keyPlaceholders: Partial<Record<AiProvider, string>> = {
        openai: 'sk-...',
        anthropic: 'sk-ant-...',
        'google-gemini': 'AIza...',
        deepseek: 'sk-...',
        groq: 'gsk_...',
        xai: 'xai-...',
        mistral: 'sk-...',
        openrouter: 'sk-or-...',
      };

      new Setting(containerEl)
        .setName('API key')
        .setDesc(keyDesc)
        .addText((text) => {
          text
            .setPlaceholder(isLocal ? '(not required)' : keyPlaceholders[provider] ?? 'sk-...')
            .setValue(this.plugin.settings.apiKey)
            .onChange(async (value) => {
              this.plugin.settings.apiKey = value.trim();
              await this.plugin.saveSettings();
            });
          text.inputEl.type = 'password';
        });

      // ── Model ──
      const modelDefault = PROVIDER_DEFAULTS[provider].model;
      new Setting(containerEl)
        .setName('Model')
        .setDesc(`Model identifier. Default for this provider: ${modelDefault || '(none)'}`)
        .addText((text) =>
          text
            .setPlaceholder(modelDefault)
            .setValue(this.plugin.settings.model)
            .onChange(async (value) => {
              this.plugin.settings.model = value.trim();
              await this.plugin.saveSettings();
            }),
        );

      // ── Base URL ──
      new Setting(containerEl)
        .setName('API base URL')
        .setDesc('Override the endpoint URL. Change only if you know what you\'re doing.')
        .addText((text) =>
          text
            .setPlaceholder(PROVIDER_DEFAULTS[provider].baseUrl)
            .setValue(this.plugin.settings.baseUrl)
            .onChange(async (value) => {
              this.plugin.settings.baseUrl = value.trim();
              await this.plugin.saveSettings();
            }),
        );

      // ── Max context ──
      new Setting(containerEl)
        .setName('Max context characters')
        .setDesc('Maximum characters of document text sent as context. Large documents will be truncated.')
        .addText((text) =>
          text
            .setPlaceholder('80000')
            .setValue(String(this.plugin.settings.maxContextChars))
            .onChange(async (value) => {
              const n = parseInt(value, 10);
              if (!isNaN(n) && n > 0) {
                this.plugin.settings.maxContextChars = n;
                await this.plugin.saveSettings();
              }
            }),
        );

      // ── Proxy auto-start (only relevant for local-proxy provider) ──
      if (provider === 'local-proxy') {
        new Setting(containerEl)
          .setName('Auto-start local proxy')
          .setDesc(
            'Automatically start claude-max-api-proxy when the plugin loads. ' +
            'Requires claude-max-api to be installed globally (npm install -g claude-max-api-proxy).',
          )
          .addToggle((toggle) =>
            toggle
              .setValue(this.plugin.settings.proxyAutoStart)
              .onChange(async (value) => {
                this.plugin.settings.proxyAutoStart = value;
                await this.plugin.saveSettings();
              }),
          );
      }

      // ── System prompt ──
      new Setting(containerEl)
        .setName('System prompt')
        .setDesc('Instructions sent to the AI before every conversation.')
        .addTextArea((text) => {
          text
            .setPlaceholder('You are a helpful assistant...')
            .setValue(this.plugin.settings.systemPrompt)
            .onChange(async (value) => {
              this.plugin.settings.systemPrompt = value;
              await this.plugin.saveSettings();
            });
          text.inputEl.rows = 5;
          text.inputEl.addClass('pcai-settings-textarea-wide');
        });
    }

    // ── PDF Reader settings ──
    new Setting(containerEl).setName('PDF reader').setHeading();

    new Setting(containerEl)
      .setName('Default zoom level')
      .setDesc('Initial zoom scale when opening a PDF (0.5 – 4.0).')
      .addText((text) =>
        text
          .setPlaceholder('1.5')
          .setValue(String(this.plugin.settings.defaultZoom))
          .onChange(async (value) => {
            const n = parseFloat(value);
            if (!isNaN(n) && n >= 0.5 && n <= 4.0) {
              this.plugin.settings.defaultZoom = n;
              await this.plugin.saveSettings();
            }
          }),
      );

    new Setting(containerEl)
      .setName('Default highlight color')
      .setDesc('Pre-selected color when highlighting text. You can still pick a different color each time.')
      .addDropdown((dd) => {
        for (const color of HIGHLIGHT_COLORS) {
          const label = this.plugin.settings.colorLabels?.[color] ?? DEFAULT_COLOR_LABELS[color];
          dd.addOption(color, `${color.charAt(0).toUpperCase() + color.slice(1)} (${label})`);
        }
        dd.setValue(this.plugin.settings.defaultHighlightColor);
        dd.onChange(async (value) => {
          this.plugin.settings.defaultHighlightColor = value as HighlightColor;
          await this.plugin.saveSettings();
        });
      });

    new Setting(containerEl)
      .setName('Resume reading position')
      .setDesc('When opening a PDF, automatically scroll to the last page you were reading.')
      .addToggle((toggle) =>
        toggle
          .setValue(this.plugin.settings.resumeLastPage)
          .onChange(async (value) => {
            this.plugin.settings.resumeLastPage = value;
            await this.plugin.saveSettings();
          }),
      );

    new Setting(containerEl)
      .setName('Annotation export format')
      .setDesc('Format used when exporting annotations to a Markdown note.')
      .addDropdown((dd) => {
        dd.addOption('callout', 'Callout blocks (> [!color])');
        dd.addOption('plain', 'Plain Markdown (quotes + bold labels)');
        dd.setValue(this.plugin.settings.exportFormat);
        dd.onChange(async (value) => {
          this.plugin.settings.exportFormat = value as ExportFormat;
          await this.plugin.saveSettings();
        });
      });

    new Setting(containerEl)
      .setName('Dictionary source')
      .setDesc('Where to look up word definitions. "auto" tries the local wordnet dictionary first, then falls back to the online API.')
      .addDropdown((dd) => {
        dd.addOption('auto', 'Auto (local + API fallback)');
        dd.addOption('local', 'Local dictionary only');
        dd.addOption('api', 'Online API only');
        dd.setValue(this.plugin.settings.dictionarySource);
        dd.onChange(async (value) => {
          this.plugin.settings.dictionarySource = value as DictionarySource;
          await this.plugin.saveSettings();
        });
      });

    // ── Highlight color labels ──
    new Setting(containerEl)
      .setName('Highlight color labels')
      .setDesc('Assign a meaning to each highlight color. These labels appear in the annotations sidebar and can be used as filters.')
      .setHeading();

    const labels = this.plugin.settings.colorLabels ?? { ...DEFAULT_COLOR_LABELS };
    for (const color of HIGHLIGHT_COLORS) {
      const setting = new Setting(containerEl)
        .setName(color.charAt(0).toUpperCase() + color.slice(1))
        .addText((text) =>
          text
            .setPlaceholder(DEFAULT_COLOR_LABELS[color])
            .setValue(labels[color] ?? DEFAULT_COLOR_LABELS[color])
            .onChange(async (value) => {
              this.plugin.settings.colorLabels[color] = value.trim() || DEFAULT_COLOR_LABELS[color];
              await this.plugin.saveSettings();
            }),
        );
      // Add a color swatch before the setting name
      const nameEl = setting.nameEl;
      const swatch = createSpan({ cls: 'pcai-settings-color-swatch' });
      swatch.setCssProps({ '--swatch-color': COLOR_HEX[color] });
      nameEl.prepend(swatch);
    }
  }
}
