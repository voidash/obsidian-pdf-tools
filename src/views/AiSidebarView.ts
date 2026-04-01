import { ItemView, WorkspaceLeaf, MarkdownRenderer, TFile } from 'obsidian';
import type PdfCanvasAiPlugin from '../main';
import type { ChatMessage, StreamOptions } from '../services/aiService';
import { VAULT_TOOLS, VaultToolExecutor } from '../services/vaultTools';
import type { Conversation } from '../store/chatStore';

export const AI_SIDEBAR_VIEW_TYPE = 'pdf-tools-sidebar';

/** Legacy scope values accepted for backward compat; context is always auto now. */
export type ContextScope = string;

interface AttachedFile {
  file: TFile;
}

/**
 * How many conversation messages to keep as "recent" when building the API
 * request. Older messages are either compacted into a summary or dropped.
 */
const RECENT_MSG_WINDOW = 10;

/**
 * Trigger compaction when the conversation exceeds this many messages and
 * the existing summary (if any) covers less than half of them.
 */
const COMPACTION_THRESHOLD = 24;

export class AiSidebarView extends ItemView {
  protected plugin: PdfCanvasAiPlugin;

  // DOM refs
  private messagesEl!: HTMLElement;
  private inputEl!: HTMLTextAreaElement;
  private sendBtn!: HTMLButtonElement;
  private contextLabelEl!: HTMLElement;
  private attachmentsEl!: HTMLElement;
  private chatListEl!: HTMLElement;

  // @ mention dropdown
  private mentionDropdownEl!: HTMLElement;
  private mentionQuery = '';
  private mentionStart = -1;

  // State
  private isStreaming = false;
  private closed = false;
  private currentFile: TFile | null = null;
  private attachedFiles: AttachedFile[] = [];
  private toolExecutor!: VaultToolExecutor;

  // Active conversation
  private activeConversation: Conversation | null = null;

  constructor(leaf: WorkspaceLeaf, plugin: PdfCanvasAiPlugin) {
    super(leaf);
    this.plugin = plugin;
  }

  getViewType(): string {
    return AI_SIDEBAR_VIEW_TYPE;
  }

  getDisplayText(): string {
    return 'PDF tools';
  }

  getIcon(): string {
    return 'bot';
  }

  async onOpen(): Promise<void> {
    await Promise.resolve();
    this.toolExecutor = new VaultToolExecutor(
      this.app,
      this.plugin.pdfService,
      () => this.plugin.getActiveCanvas(),
    );
    this.buildUI();
    this.startContextTracking();
    this.loadActiveConversation();
  }

  async onClose(): Promise<void> {
    await Promise.resolve();
    this.closed = true;
  }

  // ─── Public API ────────────────────────────────────────────────────────────

  prefillQuestion(text: string): void {
    this.inputEl.value = text;
    this.inputEl.focus();
  }

  setContextScope(_scope: ContextScope): void {
    // Kept for backward compatibility — context is always auto now
  }

  setCurrentPdf(file: TFile | null): void {
    this.currentFile = file;
    this.updateContextLabel();
  }

  // ─── Active file / canvas tracking ─────────────────────────────────────────

  private startContextTracking(): void {
    this.plugin.registerEvent(
      this.app.workspace.on('active-leaf-change', (leaf) => {
        if (leaf && (leaf.view as unknown) !== this) {
          this.onActiveLeafChange(leaf);
        }
      }),
    );
    // Run once now — find the best non-sidebar leaf
    this.onActiveLeafChange();
  }

  private onActiveLeafChange(leaf?: WorkspaceLeaf): void {
    // If no leaf provided (initial call), search for the best candidate
    if (!leaf) {
      const candidates = [
        ...this.app.workspace.getLeavesOfType('canvas'),
        ...this.app.workspace.getLeavesOfType('pdf-tools-viewer'),
        ...this.app.workspace.getLeavesOfType('markdown'),
      ];
      if (candidates.length > 0) {
        leaf = candidates[0];
      }
      if (!leaf) {
        this.updateContextLabel();
        return;
      }
    }

    const view = leaf.view;
    const viewType = view.getViewType();

    if (viewType === 'canvas') {
      this.currentFile = null;
      this.updateContextLabel();
      return;
    }

    // Check if it's our PDF viewer
    if (viewType === 'pdf-tools-viewer') {
      const pdfView = view as ItemView & { getCurrentFile?: () => unknown };
      if (typeof pdfView.getCurrentFile === 'function') {
        const file = pdfView.getCurrentFile();
        if (file instanceof TFile) {
          this.currentFile = file;
          this.updateContextLabel();
          return;
        }
      }
    }

    // Check if it's Obsidian's native PDF view, markdown view, or any file view
    const file = (view as ItemView & { file?: unknown }).file;
    if (file instanceof TFile) {
      this.currentFile = file;
      this.updateContextLabel();
      return;
    }

    this.updateContextLabel();
  }

  private updateContextLabel(): void {
    if (!this.contextLabelEl) return;

    const ctx = this.describeCurrentContext();
    this.contextLabelEl.setText(ctx);
    this.contextLabelEl.title = ctx;
  }

  private describeCurrentContext(): string {
    // Check canvas first
    const canvas = this.plugin.getActiveCanvas();
    if (canvas) {
      const canvasLeaves = this.app.workspace.getLeavesOfType('canvas');
      for (const leaf of canvasLeaves) {
        const v = leaf.view as ItemView & { canvas?: unknown; file?: { name?: string } };
        if (v?.canvas === canvas && v?.file?.name) {
          if (this.currentFile) {
            return `${v.file.name} \u203A ${this.currentFile.basename}`;
          }
          return v.file.name;
        }
      }
      if (this.currentFile) {
        return `Canvas \u203A ${this.currentFile.basename}`;
      }
      return 'Canvas';
    }

    if (this.currentFile) {
      return this.currentFile.basename;
    }

    const viewerFile = this.plugin.getViewerCurrentFile();
    if (viewerFile) {
      return viewerFile.basename;
    }

    return 'No active document';
  }

  // ─── Conversation persistence ─────────────────────────────────────────────

  private loadActiveConversation(): void {
    const store = this.plugin.chatStore;
    const activeId = store.getActiveId();

    if (activeId) {
      const conv = store.get(activeId);
      if (conv) {
        this.activeConversation = conv;
        this.replayConversation(conv);
        return;
      }
    }

    // No active conversation — create one
    this.activeConversation = store.create();
  }

  private replayConversation(conv: Conversation): void {
    this.messagesEl.empty();
    for (const msg of conv.messages) {
      if (msg.role === 'user' || msg.role === 'assistant') {
        this.renderMessage(msg.role, msg.content);
      }
    }
  }

  private ensureConversation(): Conversation {
    if (!this.activeConversation) {
      this.activeConversation = this.plugin.chatStore.create();
    }
    return this.activeConversation;
  }

  private newConversation(): void {
    this.activeConversation = this.plugin.chatStore.create();
    this.messagesEl.empty();
    this.hideChatList();
    this.inputEl.focus();
  }

  private switchConversation(id: string): void {
    const store = this.plugin.chatStore;
    const conv = store.get(id);
    if (!conv) return;

    store.setActiveId(id);
    this.activeConversation = conv;
    this.messagesEl.empty();
    this.replayConversation(conv);
    this.hideChatList();
  }

  private deleteConversation(id: string): void {
    const store = this.plugin.chatStore;
    store.delete(id);

    if (this.activeConversation?.id === id) {
      const remaining = store.getAll();
      if (remaining.length > 0) {
        this.switchConversation(remaining[0].id);
      } else {
        this.newConversation();
      }
    }

    this.renderChatList();
  }

  // ─── Chat list UI ─────────────────────────────────────────────────────────

  private toggleChatList(): void {
    if (this.chatListEl.hasClass('pcai-hidden')) {
      this.renderChatList();
      this.chatListEl.removeClass('pcai-hidden');
    } else {
      this.hideChatList();
    }
  }

  private hideChatList(): void {
    this.chatListEl.addClass('pcai-hidden');
  }

  private renderChatList(): void {
    this.chatListEl.empty();
    const conversations = this.plugin.chatStore.getAll();

    if (conversations.length === 0) {
      this.chatListEl.createDiv({ cls: 'pcai-chat-list-empty', text: 'No conversations yet' });
      return;
    }

    for (const conv of conversations) {
      const item = this.chatListEl.createDiv({
        cls: `pcai-chat-list-item${conv.id === this.activeConversation?.id ? ' pcai-chat-list-item-active' : ''}`,
      });

      const titleEl = item.createDiv({ cls: 'pcai-chat-list-title' });
      titleEl.setText(conv.title);

      const meta = item.createDiv({ cls: 'pcai-chat-list-meta' });
      const msgCount = conv.messages.filter((m) => m.role !== 'system').length;
      const dateStr = new Date(conv.updatedAt).toLocaleDateString(undefined, {
        month: 'short',
        day: 'numeric',
      });
      meta.setText(`${msgCount} msgs \u00B7 ${dateStr}`);

      const deleteBtn = item.createEl('button', {
        cls: 'pcai-chat-list-delete',
        attr: { title: 'Delete conversation' },
      });
      deleteBtn.setText('\u00D7');
      deleteBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        this.deleteConversation(conv.id);
      });

      item.addEventListener('click', () => {
        this.switchConversation(conv.id);
      });
    }
  }

  // ─── UI construction ───────────────────────────────────────────────────────

  private buildUI(): void {
    const root = this.containerEl.children[1] as HTMLElement;
    root.empty();
    root.addClass('pcai-sidebar');
    if (this.getViewType() !== AI_SIDEBAR_VIEW_TYPE) {
      root.addClass('pcai-chat-wide');
    }

    // ── Header ──
    const header = root.createDiv('pcai-sidebar-header');
    header.createSpan({ cls: 'pcai-sidebar-title', text: 'AI' });

    const btnGroup = header.createDiv('pcai-header-btns');

    const newBtn = btnGroup.createEl('button', {
      cls: 'pcai-icon-btn',
      attr: { title: 'New conversation' },
    });
    newBtn.setText('\u2795');
    newBtn.addEventListener('click', () => this.newConversation());

    const historyBtn = btnGroup.createEl('button', {
      cls: 'pcai-icon-btn',
      attr: { title: 'Chat history' },
    });
    historyBtn.setText('\uD83D\uDCAC');
    historyBtn.addEventListener('click', () => this.toggleChatList());

    // ── Chat list (hidden by default) ──
    this.chatListEl = root.createDiv('pcai-chat-list');
    this.chatListEl.addClass('pcai-hidden');

    // ── Context indicator (auto-detected) ──
    const ctxBar = root.createDiv('pcai-context-bar');
    ctxBar.createSpan({ cls: 'pcai-context-label-prefix', text: 'Context:' });
    this.contextLabelEl = ctxBar.createSpan({ cls: 'pcai-context-active' });
    this.updateContextLabel();

    // ── Messages area ──
    this.messagesEl = root.createDiv('pcai-messages');

    // ── Attachments bar (shown when files are @-mentioned) ──
    this.attachmentsEl = root.createDiv('pcai-attachments');
    this.attachmentsEl.addClass('pcai-hidden');

    // ── Input area ──
    const inputArea = root.createDiv('pcai-input-area');

    // @ mention dropdown (hidden by default)
    this.mentionDropdownEl = inputArea.createDiv('pcai-mention-dropdown');
    this.mentionDropdownEl.addClass('pcai-hidden');

    this.inputEl = inputArea.createEl('textarea', {
      cls: 'pcai-input',
      attr: { placeholder: 'Ask about your documents\u2026 use @ to attach files', rows: '3' },
    });

    this.inputEl.addEventListener('keydown', (e: KeyboardEvent) => {
      // Handle mention dropdown navigation
      if (!this.mentionDropdownEl.hasClass('pcai-hidden')) {
        if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
          e.preventDefault();
          this.navigateMentionDropdown(e.key === 'ArrowDown' ? 1 : -1);
          return;
        }
        if (e.key === 'Enter' || e.key === 'Tab') {
          const active = this.mentionDropdownEl.querySelector('.pcai-mention-item-active');
          if (active) {
            e.preventDefault();
            (active as HTMLElement).click();
            return;
          }
        }
        if (e.key === 'Escape') {
          e.preventDefault();
          this.hideMentionDropdown();
          return;
        }
      }

      if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        void this.handleSend();
      }
    });

    this.inputEl.addEventListener('input', () => this.onInputChange());

    const btnRow = inputArea.createDiv('pcai-btn-row');

    this.sendBtn = btnRow.createEl('button', {
      cls: 'pcai-send-btn mod-cta',
      text: 'Send',
    });
    this.sendBtn.addEventListener('click', () => { void this.handleSend(); });
  }

  // ─── @ mention system ──────────────────────────────────────────────────────

  private onInputChange(): void {
    const text = this.inputEl.value;
    const cursor = this.inputEl.selectionStart;

    const beforeCursor = text.slice(0, cursor);
    const atIdx = beforeCursor.lastIndexOf('@');

    if (atIdx === -1 || (atIdx > 0 && beforeCursor[atIdx - 1] !== ' ' && beforeCursor[atIdx - 1] !== '\n')) {
      this.hideMentionDropdown();
      return;
    }

    const query = beforeCursor.slice(atIdx + 1);

    if (query.includes('\n')) {
      this.hideMentionDropdown();
      return;
    }

    this.mentionStart = atIdx;
    this.mentionQuery = query;
    this.showMentionDropdown(query);
  }

  private showMentionDropdown(query: string): void {
    const queryLower = query.toLowerCase();
    const allFiles = this.app.vault.getFiles();

    const matches = allFiles
      .filter((f) => {
        if (queryLower === '') return f.extension === 'pdf' || f.extension === 'md';
        return f.path.toLowerCase().includes(queryLower) || f.basename.toLowerCase().includes(queryLower);
      })
      .sort((a, b) => {
        if (a.extension === 'pdf' && b.extension !== 'pdf') return -1;
        if (a.extension !== 'pdf' && b.extension === 'pdf') return 1;
        return a.basename.localeCompare(b.basename);
      })
      .slice(0, 10);

    this.mentionDropdownEl.empty();

    if (matches.length === 0) {
      this.mentionDropdownEl.addClass('pcai-hidden');
      return;
    }

    for (let i = 0; i < matches.length; i++) {
      const file = matches[i];
      const item = this.mentionDropdownEl.createDiv({
        cls: `pcai-mention-item${i === 0 ? ' pcai-mention-item-active' : ''}`,
      });

      const icon = file.extension === 'pdf' ? '\uD83D\uDCC4' : '\uD83D\uDCDD';
      item.createSpan({ cls: 'pcai-mention-icon', text: icon });
      item.createSpan({ cls: 'pcai-mention-name', text: file.basename });
      item.createSpan({ cls: 'pcai-mention-path', text: file.parent?.path ?? '' });

      item.addEventListener('click', () => {
        this.selectMention(file);
      });

      item.addEventListener('mouseenter', () => {
        this.mentionDropdownEl.querySelectorAll('.pcai-mention-item-active').forEach((el) =>
          el.removeClass('pcai-mention-item-active'),
        );
        item.addClass('pcai-mention-item-active');
      });
    }

    this.mentionDropdownEl.removeClass('pcai-hidden');
  }

  private hideMentionDropdown(): void {
    this.mentionDropdownEl.addClass('pcai-hidden');
    this.mentionStart = -1;
    this.mentionQuery = '';
  }

  private navigateMentionDropdown(direction: number): void {
    const items = Array.from(this.mentionDropdownEl.querySelectorAll('.pcai-mention-item'));
    const activeIdx = items.findIndex((el) => el.hasClass('pcai-mention-item-active'));
    if (activeIdx === -1) return;

    items[activeIdx].removeClass('pcai-mention-item-active');
    const newIdx = Math.max(0, Math.min(items.length - 1, activeIdx + direction));
    items[newIdx].addClass('pcai-mention-item-active');
    items[newIdx].scrollIntoView({ block: 'nearest' });
  }

  private selectMention(file: TFile): void {
    const text = this.inputEl.value;
    const before = text.slice(0, this.mentionStart);
    const after = text.slice(this.mentionStart + 1 + this.mentionQuery.length);
    this.inputEl.value = before + after;
    this.inputEl.selectionStart = this.inputEl.selectionEnd = before.length;

    if (!this.attachedFiles.some((a) => a.file.path === file.path)) {
      this.attachedFiles.push({ file });
      this.renderAttachments();
    }

    this.hideMentionDropdown();
    this.inputEl.focus();
  }

  private renderAttachments(): void {
    this.attachmentsEl.empty();

    if (this.attachedFiles.length === 0) {
      this.attachmentsEl.addClass('pcai-hidden');
      return;
    }

    this.attachmentsEl.removeClass('pcai-hidden');

    for (const af of this.attachedFiles) {
      const chip = this.attachmentsEl.createDiv('pcai-attach-chip');
      const icon = af.file.extension === 'pdf' ? '\uD83D\uDCC4' : '\uD83D\uDCDD';
      chip.createSpan({ text: `${icon} ${af.file.basename}` });

      const removeBtn = chip.createSpan({ cls: 'pcai-attach-remove', text: '\u00D7' });
      removeBtn.addEventListener('click', () => {
        this.attachedFiles = this.attachedFiles.filter((a) => a.file.path !== af.file.path);
        this.renderAttachments();
      });
    }
  }

  // ─── Message handling ──────────────────────────────────────────────────────

  private async handleSend(): Promise<void> {
    if (this.isStreaming) return;

    const userText = this.inputEl.value.trim();
    if (!userText) return;

    this.setStreaming(true);
    this.inputEl.value = '';
    this.hideMentionDropdown();

    this.renderMessage('user', userText);

    const conv = this.ensureConversation();

    // ── Build context ──
    const contextParts: string[] = [];

    // Auto context: active file (PDF or markdown)
    const activeFile = this.currentFile ?? this.plugin.getViewerCurrentFile();
    if (activeFile) {
      try {
        if (activeFile.extension === 'pdf') {
          const text = await this.plugin.pdfService.extractText(activeFile);
          contextParts.push(`[Active PDF: ${activeFile.name}]\n${text}`);
        } else {
          const content = await this.app.vault.cachedRead(activeFile);
          contextParts.push(`[Active file: ${activeFile.name}]\n${content}`);
        }
      } catch (err) {
        console.error('PDF Tools \u2014 auto context error:', err);
      }
    }

    // Auto context: active canvas
    const canvas = this.plugin.getActiveCanvas();
    if (canvas) {
      try {
        const canvasText = await this.plugin.gatherCanvasContext();
        if (!canvasText.startsWith('[No active')) {
          contextParts.push(canvasText);
        }
      } catch {
        // Skip
      }
    }

    // @-attached files
    for (const af of this.attachedFiles) {
      try {
        if (af.file.extension === 'pdf') {
          const text = await this.plugin.pdfService.extractText(af.file);
          contextParts.push(`[Attached: ${af.file.name}]\n${text}`);
        } else {
          const content = await this.app.vault.cachedRead(af.file);
          contextParts.push(`[Attached: ${af.file.name}]\n${content}`);
        }
      } catch {
        contextParts.push(`[Attached: ${af.file.name}] (failed to read)`);
      }
    }

    this.attachedFiles = [];
    this.renderAttachments();

    if (this.closed) {
      this.setStreaming(false);
      return;
    }

    // ── System prompt ──
    let systemContent = this.plugin.settings.systemPrompt;
    systemContent +=
      '\n\nYou have access to tools to search and read files in the user\'s Obsidian vault. ' +
      'Use them when you need to look up additional information, find related files, or explore the vault. ' +
      'The user may attach files with @mentions \u2014 their content is provided in the context.';

    const currentCtx = this.describeCurrentContext();
    if (currentCtx !== 'No active document') {
      systemContent += `\n\nThe user is currently looking at: ${currentCtx}`;
    }

    // ── Build API messages using compaction + sliding window ──
    const apiMessages: ChatMessage[] = [
      { role: 'system', content: systemContent },
    ];

    // If we have a compacted summary, include it as context
    if (conv.compactedSummary) {
      apiMessages.push({
        role: 'system',
        content: `Summary of earlier conversation:\n\n${conv.compactedSummary}`,
      });
    }

    // Sliding window: take the last N messages (after the compacted ones)
    const startIdx = conv.compactedCount ?? 0;
    const recentMessages = conv.messages.slice(startIdx);
    const windowStart = Math.max(0, recentMessages.length - RECENT_MSG_WINDOW);
    for (let i = windowStart; i < recentMessages.length; i++) {
      apiMessages.push(recentMessages[i]);
    }

    // Build user content with context
    let userContent = userText;
    if (contextParts.length > 0) {
      let combined = contextParts.join('\n\n---\n\n');
      if (combined.length > this.plugin.settings.maxContextChars) {
        combined = combined.slice(0, this.plugin.settings.maxContextChars) + '\n\n[\u2026truncated]';
      }
      userContent = `Context:\n\n${combined}\n\n---\n\nQuestion: ${userText}`;
    }

    // Save user message to store (store the raw text, not the context-enriched version)
    this.plugin.chatStore.addMessage(conv.id, { role: 'user', content: userText });

    // Ensure local proxy is running before making the request
    if (this.plugin.settings.provider === 'local-proxy' && this.plugin.proxyManager) {
      await this.plugin.proxyManager.ensureRunning();
    }

    const streamEl = this.createStreamingMessage();
    let accumulated = '';

    const streamOptions: StreamOptions = {
      tools: VAULT_TOOLS,
      onToolCall: (event) => {
        this.showToolCallStatus(event.name, event.args);
      },
      executeToolCall: (name, args) => {
        return this.toolExecutor.execute(name, args);
      },
    };

    await this.plugin.aiService.streamChat(
      [...apiMessages, { role: 'user', content: userContent }],
      (delta) => {
        if (this.closed) return;
        accumulated += delta;
        this.updateStreamingContent(streamEl, accumulated);
      },
      () => {
        if (this.closed) return;
        this.finalizeStreamingMessage(streamEl, accumulated);

        // Save assistant message to store
        this.plugin.chatStore.addMessage(conv.id, { role: 'assistant', content: accumulated });

        this.setStreaming(false);

        // Check if compaction is needed (background, non-blocking)
        void this.maybeCompact(conv).catch((e: unknown) => {
          console.error('PDF tools: compaction error', e);
        });
      },
      (errorMsg) => {
        if (this.closed) return;
        this.removeCursor(streamEl);
        this.renderError(errorMsg);
        this.setStreaming(false);
      },
      streamOptions,
    );
  }

  // ─── Compaction ───────────────────────────────────────────────────────────

  /**
   * If the conversation is long enough and the existing summary is stale,
   * ask the AI to summarize the older portion so we can keep the API
   * context window manageable.
   */
  private async maybeCompact(conv: Conversation): Promise<void> {
    const totalMessages = conv.messages.length;
    if (totalMessages < COMPACTION_THRESHOLD) return;

    const alreadyCompacted = conv.compactedCount ?? 0;
    const uncompacted = totalMessages - alreadyCompacted;

    // Only compact if there are enough uncompacted messages beyond our window
    if (uncompacted <= RECENT_MSG_WINDOW + 4) return;

    // Messages to summarize: everything except the last RECENT_MSG_WINDOW
    const toSummarize = conv.messages.slice(0, totalMessages - RECENT_MSG_WINDOW);

    // Build a transcript for the summarizer
    const transcript = toSummarize
      .filter((m) => m.role !== 'system')
      .map((m) => `${m.role === 'user' ? 'User' : 'Assistant'}: ${m.content}`)
      .join('\n\n');

    // Include the old summary if there is one, so it gets folded in
    const oldSummary = conv.compactedSummary
      ? `Previous summary:\n${conv.compactedSummary}\n\n`
      : '';

    const summaryPrompt: ChatMessage[] = [
      {
        role: 'system',
        content:
          'You are a conversation summarizer. Produce a concise but thorough summary of the ' +
          'conversation below. Preserve all key facts, decisions, file names, code snippets, ' +
          'and important context. The summary will replace these messages in future API calls, ' +
          'so nothing important should be lost. Output only the summary, no preamble.',
      },
      {
        role: 'user',
        content: `${oldSummary}Conversation to summarize:\n\n${transcript}`,
      },
    ];

    // Collect the summary via streaming (we just accumulate, no UI)
    let summary = '';
    await new Promise<void>((resolve) => {
      void this.plugin.aiService.streamChat(
        summaryPrompt,
        (delta) => {
          summary += delta;
        },
        () => resolve(),
        (err) => {
          console.error('PDF Tools: compaction summarization failed:', err);
          resolve();
        },
      );
    });

    if (summary.length > 0) {
      const newCompactedCount = totalMessages - RECENT_MSG_WINDOW;
      this.plugin.chatStore.setCompaction(conv.id, summary, newCompactedCount);
    }
  }

  // ─── Render helpers ────────────────────────────────────────────────────────

  private renderMessage(role: 'user' | 'assistant', content: string): HTMLElement {
    const el = this.messagesEl.createDiv({ cls: `pcai-msg pcai-msg-${role}` });
    if (role === 'assistant') {
      this.renderMarkdown(el, content);
    } else {
      el.createDiv({ cls: 'pcai-msg-content', text: content });
    }
    this.scrollToBottom();
    return el;
  }

  private showToolCallStatus(toolName: string, argsJson: string): void {
    let label = toolName;
    try {
      const args = JSON.parse(argsJson) as Record<string, unknown>;
      if (toolName === 'search_vault' && typeof args.query === 'string') {
        label = `Searching vault for \u201C${args.query}\u201D`;
      } else if (toolName === 'read_file' && typeof args.path === 'string') {
        label = `Reading ${args.path}`;
      } else if (toolName === 'list_files') {
        label = 'Listing vault files';
      } else if (toolName === 'get_canvas_items') {
        label = 'Reading canvas items';
      }
    } catch {
      // Use raw name
    }

    const el = this.messagesEl.createDiv({ cls: 'pcai-tool-status' });
    el.createSpan({ cls: 'pcai-tool-icon', text: '\uD83D\uDD0D' });
    el.createSpan({ text: label });
    this.scrollToBottomIfNeeded();
  }

  private createStreamingMessage(): HTMLElement {
    const el = this.messagesEl.createDiv({ cls: 'pcai-msg pcai-msg-assistant pcai-msg-streaming' });
    const content = el.createDiv({ cls: 'pcai-msg-content' });
    content.createSpan({ cls: 'pcai-cursor' });
    this.scrollToBottom();
    return el;
  }

  private updateStreamingContent(msgEl: HTMLElement, fullText: string): void {
    const content = msgEl.querySelector('.pcai-msg-content') as HTMLElement;
    if (!content) return;

    content.empty();
    this.renderMarkdown(content, fullText);
    content.createSpan({ cls: 'pcai-cursor' });
    this.scrollToBottomIfNeeded();
  }

  private finalizeStreamingMessage(msgEl: HTMLElement, fullText: string): void {
    msgEl.removeClass('pcai-msg-streaming');
    const content = msgEl.querySelector('.pcai-msg-content') as HTMLElement;
    if (!content) return;

    content.empty();
    this.renderMarkdown(content, fullText);
  }

  private removeCursor(msgEl: HTMLElement): void {
    msgEl.removeClass('pcai-msg-streaming');
    const cursor = msgEl.querySelector('.pcai-cursor');
    if (cursor) cursor.remove();
  }

  private renderMarkdown(containerEl: HTMLElement, markdown: string): void {
    const wrapper = containerEl.createDiv({ cls: 'pcai-msg-content pcai-markdown' });
    try {
      void MarkdownRenderer.render(
        this.app,
        markdown,
        wrapper,
        '',
        this,
      );
    } catch {
      wrapper.setText(markdown);
    }
  }

  private renderError(text: string): void {
    const el = this.messagesEl.createDiv({ cls: 'pcai-msg pcai-msg-error' });
    el.createDiv({ cls: 'pcai-msg-content', text });
    this.scrollToBottom();
  }

  private setStreaming(active: boolean): void {
    this.isStreaming = active;
    this.sendBtn.disabled = active;
    this.inputEl.disabled = active;
    if (!active && !this.closed) this.inputEl.focus();
  }

  /** Check if the user is scrolled near the bottom (within 80px). */
  private isNearBottom(): boolean {
    const el = this.messagesEl;
    return el.scrollHeight - el.scrollTop - el.clientHeight < 80;
  }

  /** Always scroll to the bottom (e.g. after user sends a message). */
  private scrollToBottom(): void {
    this.messagesEl.scrollTop = this.messagesEl.scrollHeight;
  }

  /** Scroll to bottom only if the user hasn't scrolled up to read earlier content. */
  private scrollToBottomIfNeeded(): void {
    if (this.isNearBottom()) {
      this.scrollToBottom();
    }
  }
}
