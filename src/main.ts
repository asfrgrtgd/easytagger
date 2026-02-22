import {
  ButtonComponent,
  ItemView,
  Notice,
  Plugin,
  Setting,
  TFile,
  WorkspaceLeaf,
  requestUrl,
} from "obsidian";

/**
 * Update these constants for your environment.
 */
const LMSTUDIO_BASE_URL = "http://192.168.10.105:1234";
const LMSTUDIO_MODEL = "gpt-oss-swallow-20b-rl-v0.1";
const TAG_LIST_PATH = "_system/allowed-tags.md";

const MAX_GENERATED_TAGS = 5;
const MIN_NEW_SUGGESTIONS = 2;
const MAX_NEW_SUGGESTIONS = 5;
const MAX_NOTE_CHARS_FOR_PROMPT = 8000;
const RETRY_MAX = 3;
const RETRY_DELAY_MS = 250;
const PREFER_STRUCTURED_OUTPUT = true;
const RIBBON_ICON = "tags";
const VIEW_TYPE_TAG_GENERATOR = "local-lm-tag-generator-sidebar";
const SIDEBAR_VIEW_TITLE = "Tag Generator";

type NewTagMode = "none" | "approval" | "auto";
type CoverageState = "sufficient" | "insufficient";
type ReviewVerdict = "accept" | "review" | "reject";

type ChatMessage = {
  role: "system" | "user";
  content: string;
};

interface PluginSettings {
  newTagMode: NewTagMode;
  newSuggestionCount: number;
  enableSecondaryReview: boolean;
  autoAppendToAllowedList: boolean;
}

interface NewTagSuggestion {
  tag: string;
  reason: string;
  confidence: number;
}

interface ReviewedSuggestion extends NewTagSuggestion {
  granularityScore: number;
  toneScore: number;
  verdict: ReviewVerdict;
  reviewReason: string;
  selected: boolean;
}

interface TagGenerationPayload {
  existingTags: string[];
  coverage: CoverageState;
  newSuggestions: NewTagSuggestion[];
}

interface TagRunResult {
  filePath: string;
  existingTags: string[];
  existingAddedTags: string[];
  coverage: CoverageState;
  reviewedSuggestions: ReviewedSuggestion[];
  appliedNewTags: string[];
  updatedAt: number;
}

type ValidationOk<T> = { ok: true; value: T };
type ValidationNg = { ok: false; reason: string };
type ValidationResult<T> = ValidationOk<T> | ValidationNg;

const DEFAULT_SETTINGS: PluginSettings = {
  newTagMode: "approval",
  newSuggestionCount: 3,
  enableSecondaryReview: true,
  autoAppendToAllowedList: false,
};

class LmStudioRequestError extends Error {
  status?: number;
  bodyText?: string;

  constructor(message: string, status?: number, bodyText?: string) {
    super(message);
    this.name = "LmStudioRequestError";
    this.status = status;
    this.bodyText = bodyText;
  }
}

class TagGeneratorSidebarView extends ItemView {
  plugin: LocalLmTagGeneratorPlugin;

  constructor(leaf: WorkspaceLeaf, plugin: LocalLmTagGeneratorPlugin) {
    super(leaf);
    this.plugin = plugin;
  }

  getViewType(): string {
    return VIEW_TYPE_TAG_GENERATOR;
  }

  getDisplayText(): string {
    return SIDEBAR_VIEW_TITLE;
  }

  getIcon(): string {
    return RIBBON_ICON;
  }

  async onOpen(): Promise<void> {
    await this.render();
  }

  async onClose(): Promise<void> {
    this.contentEl.empty();
  }

  async render(): Promise<void> {
    const container = this.contentEl;
    container.empty();
    container.addClass("local-lm-tag-generator-view");

    container.createEl("h4", { text: "Tag generation" });
    container.createEl("p", {
      text: "Generate existing tags, then optionally create and review new tags.",
    });

    this.renderControls(container);
    this.renderActions(container);
    this.renderResult(container);
  }

  renderControls(container: HTMLElement): void {
    const settings = this.plugin.settings;

    new Setting(container)
      .setName("New tag mode")
      .setDesc("New tags only: none / approval / auto")
      .addDropdown((dropdown) => {
        dropdown.addOption("none", "なし");
        dropdown.addOption("approval", "承認式");
        dropdown.addOption("auto", "自動追加");
        dropdown.setValue(settings.newTagMode);
        dropdown.onChange(async (value) => {
          await this.plugin.updateNewTagMode(value as NewTagMode);
        });
      });

    new Setting(container)
      .setName("New suggestion count")
      .setDesc("How many new tag candidates to ask for (2-5)")
      .addDropdown((dropdown) => {
        for (let i = MIN_NEW_SUGGESTIONS; i <= MAX_NEW_SUGGESTIONS; i++) {
          dropdown.addOption(String(i), String(i));
        }
        dropdown.setValue(String(settings.newSuggestionCount));
        dropdown.onChange(async (value) => {
          const parsed = Number(value);
          await this.plugin.updateNewSuggestionCount(parsed);
        });
      });

    new Setting(container)
      .setName("Secondary review agent")
      .setDesc("Check granularity/tone similarity against existing tags")
      .addToggle((toggle) => {
        toggle.setValue(settings.enableSecondaryReview);
        toggle.onChange(async (value) => {
          await this.plugin.updateSecondaryReview(value);
        });
      });

    new Setting(container)
      .setName("Auto append new tags to allowed list")
      .setDesc("Initial value is OFF")
      .addToggle((toggle) => {
        toggle.setValue(settings.autoAppendToAllowedList);
        toggle.onChange(async (value) => {
          await this.plugin.updateAutoAppendAllowedList(value);
        });
      });
  }

  renderActions(container: HTMLElement): void {
    const actionRow = container.createDiv({ cls: "local-lm-actions" });

    const runButton = new ButtonComponent(actionRow)
      .setButtonText(this.plugin.isBusy ? "Running..." : "Run Tag generation")
      .setCta();

    if (this.plugin.isBusy) {
      runButton.setDisabled(true);
    }

    runButton.onClick(async () => {
      await this.plugin.runTagGenerationWorkflow("sidebar");
      await this.render();
    });

    if (this.plugin.settings.newTagMode === "approval") {
      const applyButton = new ButtonComponent(actionRow).setButtonText(
        "Apply selected new tags",
      );

      if (this.plugin.isBusy || !this.plugin.canApplySelectedSuggestions()) {
        applyButton.setDisabled(true);
      }

      applyButton.onClick(async () => {
        await this.plugin.applySelectedSuggestions();
        await this.render();
      });
    }
  }

  renderResult(container: HTMLElement): void {
    const result = this.plugin.lastRunResult;
    if (!result) {
      container.createEl("p", {
        cls: "local-lm-muted",
        text: "Run once to see selected tags and new tag suggestions.",
      });
      return;
    }

    const resultWrap = container.createDiv({ cls: "local-lm-result" });
    resultWrap.createEl("hr");

    const meta = resultWrap.createDiv({ cls: "local-lm-meta" });
    meta.createEl("div", { text: `File: ${result.filePath}` });
    meta.createEl("div", {
      text: `Coverage: ${result.coverage === "sufficient" ? "十分" : "不足"}`,
    });

    this.renderTagGroup(
      resultWrap,
      "Existing tags",
      result.existingTags,
      "local-lm-chip-existing",
    );

    if (result.existingAddedTags.length > 0) {
      this.renderTagGroup(
        resultWrap,
        "Added existing tags",
        result.existingAddedTags,
        "local-lm-chip-added",
      );
    }

    if (result.appliedNewTags.length > 0) {
      this.renderTagGroup(
        resultWrap,
        "Applied new tags",
        result.appliedNewTags,
        "local-lm-chip-new",
      );
    }

    if (result.reviewedSuggestions.length > 0) {
      const section = resultWrap.createDiv({ cls: "local-lm-suggestions" });
      section.createEl("h5", { text: "New tag suggestions" });

      for (const suggestion of result.reviewedSuggestions) {
        const row = section.createDiv({ cls: "local-lm-suggestion-row" });

        if (this.plugin.settings.newTagMode === "approval") {
          const checkbox = row.createEl("input", {
            type: "checkbox",
            cls: "local-lm-suggestion-checkbox",
          });
          checkbox.checked = suggestion.selected;
          checkbox.disabled = suggestion.verdict === "reject";
          checkbox.onchange = () => {
            this.plugin.setSuggestionSelected(suggestion.tag, checkbox.checked);
          };
        }

        const body = row.createDiv({ cls: "local-lm-suggestion-body" });
        const head = body.createDiv({ cls: "local-lm-suggestion-head" });
        head.createEl("code", { text: suggestion.tag });
        head.createEl("span", {
          cls: `local-lm-badge local-lm-badge-${suggestion.verdict}`,
          text: verdictLabel(suggestion.verdict),
        });

        body.createEl("div", {
          cls: "local-lm-suggestion-reason",
          text: suggestion.reason,
        });
        body.createEl("div", {
          cls: "local-lm-suggestion-scores",
          text: `confidence ${Math.round(suggestion.confidence * 100)} / granularity ${suggestion.granularityScore} / tone ${suggestion.toneScore}`,
        });
        body.createEl("div", {
          cls: "local-lm-suggestion-review",
          text: suggestion.reviewReason,
        });
      }
    } else {
      resultWrap.createEl("p", {
        cls: "local-lm-muted",
        text: "No new tag suggestions.",
      });
    }
  }

  renderTagGroup(
    container: HTMLElement,
    title: string,
    tags: string[],
    chipClass: string,
  ): void {
    if (tags.length === 0) return;

    const group = container.createDiv({ cls: "local-lm-group" });
    group.createEl("h5", { text: title });
    const chips = group.createDiv({ cls: "local-lm-chip-row" });

    for (const tag of tags) {
      chips.createEl("span", {
        cls: `local-lm-chip ${chipClass}`,
        text: tag,
      });
    }
  }
}

export default class LocalLmTagGeneratorPlugin extends Plugin {
  settings: PluginSettings = { ...DEFAULT_SETTINGS };
  lastRunResult: TagRunResult | null = null;
  isBusy = false;

  async onload(): Promise<void> {
    try {
      await this.loadSettings();

      this.registerView(
        VIEW_TYPE_TAG_GENERATOR,
        (leaf) => new TagGeneratorSidebarView(leaf, this),
      );

      this.addRibbonIcon(RIBBON_ICON, "Tag generation", async () => {
        await this.runTagGenerationWorkflow("command");
      });

      this.addCommand({
        id: "generate-tags-from-local-lmstudio",
        name: "Tag generation",
        callback: async () => {
          await this.runTagGenerationWorkflow("command");
        },
      });

      this.addCommand({
        id: "generate-tags-from-local-lmstudio-japanese",
        name: "タグ生成を実行",
        callback: async () => {
          await this.runTagGenerationWorkflow("command");
        },
      });

      this.addCommand({
        id: "open-tag-generator-sidebar",
        name: "Open Tag generator sidebar",
        callback: async () => {
          await this.activateSidebarView();
        },
      });

      this.addCommand({
        id: "apply-selected-new-tags",
        name: "Apply selected new tag suggestions",
        callback: async () => {
          await this.applySelectedSuggestions();
        },
      });

      const workspace = this.app.workspace as {
        onLayoutReady?: (callback: () => void) => void;
      };

      if (typeof workspace.onLayoutReady === "function") {
        workspace.onLayoutReady(() => {
          void this.activateSidebarView();
        });
      }

      new Notice("Local LM Tag Generator loaded");
    } catch (err) {
      console.error("[local-lm-tag-generator] onload failed", err);
      new Notice(`Local LM Tag Generator failed to load: ${errorToMessage(err)}`);
    }
  }

  async onunload(): Promise<void> {
    this.app.workspace
      .getLeavesOfType(VIEW_TYPE_TAG_GENERATOR)
      .forEach((leaf) => leaf.detach());
  }

  async activateSidebarView(): Promise<void> {
    const { workspace } = this.app;
    const existingLeaves = workspace.getLeavesOfType(VIEW_TYPE_TAG_GENERATOR);
    if (existingLeaves.length > 0) {
      if (typeof (workspace as { revealLeaf?: unknown }).revealLeaf === "function") {
        await (workspace as { revealLeaf: (leaf: WorkspaceLeaf) => Promise<void> }).revealLeaf(
          existingLeaves[0],
        );
      }
      return;
    }

    if (typeof (workspace as { getLeftLeaf?: unknown }).getLeftLeaf !== "function") {
      return;
    }

    const leaf = (
      workspace as { getLeftLeaf: (split: boolean) => WorkspaceLeaf | null }
    ).getLeftLeaf(false);
    if (!leaf) return;

    await leaf.setViewState({
      type: VIEW_TYPE_TAG_GENERATOR,
      active: false,
    });

    if (typeof (workspace as { revealLeaf?: unknown }).revealLeaf === "function") {
      await (workspace as { revealLeaf: (targetLeaf: WorkspaceLeaf) => Promise<void> }).revealLeaf(
        leaf,
      );
    }
  }

  async updateNewTagMode(mode: NewTagMode): Promise<void> {
    this.settings.newTagMode = mode;
    await this.saveSettings();
    this.refreshSidebarViews();
  }

  async updateNewSuggestionCount(value: number): Promise<void> {
    this.settings.newSuggestionCount = clamp(
      value,
      MIN_NEW_SUGGESTIONS,
      MAX_NEW_SUGGESTIONS,
    );
    await this.saveSettings();
    this.refreshSidebarViews();
  }

  async updateSecondaryReview(enabled: boolean): Promise<void> {
    this.settings.enableSecondaryReview = enabled;
    await this.saveSettings();
    this.refreshSidebarViews();
  }

  async updateAutoAppendAllowedList(enabled: boolean): Promise<void> {
    this.settings.autoAppendToAllowedList = enabled;
    await this.saveSettings();
    this.refreshSidebarViews();
  }

  canApplySelectedSuggestions(): boolean {
    if (this.settings.newTagMode !== "approval") return false;
    if (!this.lastRunResult) return false;

    return this.lastRunResult.reviewedSuggestions.some(
      (s) => s.selected && s.verdict !== "reject",
    );
  }

  setSuggestionSelected(tag: string, selected: boolean): void {
    if (!this.lastRunResult) return;
    const key = normalizeTag(tag).toLowerCase();
    for (const suggestion of this.lastRunResult.reviewedSuggestions) {
      if (suggestion.tag.toLowerCase() === key) {
        suggestion.selected = selected && suggestion.verdict !== "reject";
      }
    }
  }

  async applySelectedSuggestions(): Promise<void> {
    if (!this.lastRunResult) {
      new Notice("No pending suggestion result.");
      return;
    }

    if (this.settings.newTagMode !== "approval") {
      new Notice("New tag mode is not set to approval.");
      return;
    }

    if (this.isBusy) {
      new Notice("Tag generation is running. Please wait.");
      return;
    }

    const target = this.app.vault.getAbstractFileByPath(this.lastRunResult.filePath);
    if (!(target instanceof TFile)) {
      new Notice(`Target note not found: ${this.lastRunResult.filePath}`);
      return;
    }

    const selected = this.lastRunResult.reviewedSuggestions
      .filter((s) => s.selected && s.verdict !== "reject")
      .sort((a, b) => scoreSuggestion(b) - scoreSuggestion(a));

    const slots = Math.max(0, MAX_GENERATED_TAGS - this.lastRunResult.existingTags.length);
    const candidateTags = selected.slice(0, slots).map((s) => s.tag);

    if (candidateTags.length === 0) {
      new Notice("No selectable new tags.");
      return;
    }

    this.isBusy = true;
    this.refreshSidebarViews();

    try {
      const applied = await this.applyTagsToFrontmatter(target, candidateTags);

      if (this.settings.autoAppendToAllowedList && applied.added.length > 0) {
        await this.appendTagsToAllowedList(applied.added);
      }

      this.lastRunResult.appliedNewTags = uniqueTags([
        ...this.lastRunResult.appliedNewTags,
        ...applied.added,
      ]);
      this.lastRunResult.updatedAt = Date.now();

      if (applied.added.length > 0) {
        new Notice(`Applied new tags: ${applied.added.join(", ")}`);
      } else {
        new Notice("No additional new tags were applied.");
      }
    } catch (err) {
      console.error("[local-lm-tag-generator]", err);
      new Notice(`Failed to apply suggestions: ${errorToMessage(err)}`);
    } finally {
      this.isBusy = false;
      this.refreshSidebarViews();
    }
  }

  async runTagGenerationWorkflow(source: "sidebar" | "command"): Promise<void> {
    if (this.isBusy) {
      new Notice("Tag generation is already running.");
      return;
    }

    this.isBusy = true;
    this.refreshSidebarViews();

    try {
      const file = this.app.workspace.getActiveFile();
      if (!file || file.extension !== "md") {
        new Notice("Open a Markdown note before running this command.");
        return;
      }

      const allowedTags = await this.loadAllowedTags();
      if (allowedTags.length === 0) {
        new Notice(`No allowed tags found at: ${TAG_LIST_PATH}`);
        return;
      }

      const rawNote = await this.app.vault.read(file);
      const noteBody = stripYamlFrontmatter(rawNote);
      const noteForPrompt = trimText(noteBody, MAX_NOTE_CHARS_FOR_PROMPT);

      const plan = await this.generateTagPlan(file, noteForPrompt, allowedTags);

      const existingApplyResult = await this.applyTagsToFrontmatter(file, plan.existingTags);

      let reviewedSuggestions: ReviewedSuggestion[] = [];
      if (this.settings.newTagMode !== "none" && plan.newSuggestions.length > 0) {
        if (this.settings.enableSecondaryReview) {
          reviewedSuggestions = await this.reviewNewSuggestions(
            file,
            noteForPrompt,
            allowedTags,
            plan.existingTags,
            plan.newSuggestions,
          );
        } else {
          reviewedSuggestions = plan.newSuggestions.map((s) => ({
            ...s,
            granularityScore: 100,
            toneScore: 100,
            verdict: "accept",
            reviewReason: "Secondary review disabled",
            selected: false,
          }));
        }
      }

      let appliedNewTags: string[] = [];
      if (this.settings.newTagMode === "auto" && reviewedSuggestions.length > 0) {
        const slots = Math.max(0, MAX_GENERATED_TAGS - plan.existingTags.length);
        const ranked = reviewedSuggestions
          .filter((s) => s.verdict !== "reject")
          .sort((a, b) => scoreSuggestion(b) - scoreSuggestion(a));

        const candidateTags = ranked.slice(0, slots).map((s) => s.tag);
        if (candidateTags.length > 0) {
          const newApplyResult = await this.applyTagsToFrontmatter(file, candidateTags);
          appliedNewTags = newApplyResult.added;

          if (this.settings.autoAppendToAllowedList && appliedNewTags.length > 0) {
            await this.appendTagsToAllowedList(appliedNewTags);
          }
        }
      }

      for (const suggestion of reviewedSuggestions) {
        suggestion.selected =
          this.settings.newTagMode === "approval" && suggestion.verdict === "accept";
      }

      this.lastRunResult = {
        filePath: file.path,
        existingTags: plan.existingTags,
        existingAddedTags: existingApplyResult.added,
        coverage: plan.coverage,
        reviewedSuggestions,
        appliedNewTags,
        updatedAt: Date.now(),
      };

      if (this.settings.newTagMode === "approval" && reviewedSuggestions.length > 0) {
        await this.activateSidebarView();
      }

      this.notifyRunResult(source, this.lastRunResult);
    } catch (err) {
      console.error("[local-lm-tag-generator]", err);
      new Notice(`Tag generation failed: ${errorToMessage(err)}`);
    } finally {
      this.isBusy = false;
      this.refreshSidebarViews();
    }
  }

  notifyRunResult(source: "sidebar" | "command", result: TagRunResult): void {
    const pieces: string[] = [];

    if (result.existingAddedTags.length > 0) {
      pieces.push(`existing +${result.existingAddedTags.length}`);
    }

    if (result.appliedNewTags.length > 0) {
      pieces.push(`new +${result.appliedNewTags.length}`);
    }

    if (pieces.length === 0) {
      pieces.push("no tags added");
    }

    if (this.settings.newTagMode === "approval" && result.reviewedSuggestions.length > 0) {
      const reviewCount = result.reviewedSuggestions.filter(
        (s) => s.verdict !== "reject",
      ).length;
      new Notice(
        `Tag generation done (${pieces.join(", ")}). ${reviewCount} new suggestions are waiting for approval in sidebar.`,
      );
      return;
    }

    if (source === "command" && this.settings.newTagMode === "approval") {
      new Notice(`Tag generation done (${pieces.join(", ")}). Open sidebar to review new tags.`);
      return;
    }

    new Notice(`Tag generation done (${pieces.join(", ")}).`);
  }

  async loadAllowedTags(): Promise<string[]> {
    const abs = this.app.vault.getAbstractFileByPath(TAG_LIST_PATH);
    if (!(abs instanceof TFile)) return [];
    const raw = await this.app.vault.read(abs);
    return parseAllowedTagFile(raw);
  }

  async appendTagsToAllowedList(tags: string[]): Promise<string[]> {
    const incoming = uniqueTags(tags);
    if (incoming.length === 0) return [];

    const abs = this.app.vault.getAbstractFileByPath(TAG_LIST_PATH);
    let current: string[] = [];

    if (abs instanceof TFile) {
      const raw = await this.app.vault.read(abs);
      current = parseAllowedTagFile(raw);
    }

    const merged = uniqueTags([...current, ...incoming]);
    const currentSet = new Set(current.map((t) => t.toLowerCase()));
    const added = merged.filter((t) => !currentSet.has(t.toLowerCase()));

    if (added.length === 0) return [];

    const output = renderAllowedTagFile(merged);

    if (abs instanceof TFile) {
      await this.app.vault.modify(abs, output);
    } else {
      await this.ensureParentFolders(TAG_LIST_PATH);
      await this.app.vault.create(TAG_LIST_PATH, output);
    }

    return added;
  }

  async ensureParentFolders(filePath: string): Promise<void> {
    const parts = filePath.split("/");
    parts.pop();

    let current = "";
    for (const part of parts) {
      if (!part) continue;
      current = current ? `${current}/${part}` : part;
      const abs = this.app.vault.getAbstractFileByPath(current);
      if (abs) continue;
      await this.app.vault.createFolder(current);
    }
  }

  async generateTagPlan(
    file: TFile,
    noteText: string,
    allowedTags: string[],
  ): Promise<TagGenerationPayload> {
    let lastReason = "unknown";
    let useStructuredOutput = PREFER_STRUCTURED_OUTPUT;
    const targetSuggestionCount = this.settings.newTagMode === "none"
      ? 0
      : clamp(this.settings.newSuggestionCount, MIN_NEW_SUGGESTIONS, MAX_NEW_SUGGESTIONS);

    for (let attempt = 1; attempt <= RETRY_MAX; attempt++) {
      try {
        const content = await callLmStudioChat({
          messages: buildGenerationMessages({
            file,
            noteText,
            allowedTags,
            mode: this.settings.newTagMode,
            suggestionCount: targetSuggestionCount,
            maxExistingTags: MAX_GENERATED_TAGS,
            attempt,
          }),
          responseFormat: useStructuredOutput
            ? buildGenerationResponseFormat(allowedTags)
            : undefined,
          temperature: 0,
          maxTokens: 500,
        });

        if (!content.trim()) {
          lastReason = "empty response";
        } else {
          const parsed = extractJsonObjectStrict(content);
          if (!parsed) {
            lastReason = "JSON parse failed";
          } else {
            const checked = validateGenerationPayload(
              parsed,
              allowedTags,
              this.settings.newTagMode,
              targetSuggestionCount,
            );
            if (checked.ok) return checked.value;
            lastReason = checked.reason;
          }
        }
      } catch (err) {
        lastReason = errorToMessage(err);
        if (useStructuredOutput && looksLikeStructuredOutputUnsupported(err)) {
          useStructuredOutput = false;
        }
      }

      if (attempt < RETRY_MAX) {
        await sleep(RETRY_DELAY_MS * attempt);
      }
    }

    throw new Error(`Could not generate tags (reason: ${lastReason})`);
  }

  async reviewNewSuggestions(
    file: TFile,
    noteText: string,
    allowedTags: string[],
    existingTags: string[],
    suggestions: NewTagSuggestion[],
  ): Promise<ReviewedSuggestion[]> {
    if (suggestions.length === 0) return [];

    let lastReason = "unknown";
    let useStructuredOutput = PREFER_STRUCTURED_OUTPUT;

    for (let attempt = 1; attempt <= RETRY_MAX; attempt++) {
      try {
        const content = await callLmStudioChat({
          messages: buildReviewMessages({
            file,
            noteText,
            allowedTags,
            existingTags,
            suggestions,
            attempt,
          }),
          responseFormat: useStructuredOutput
            ? buildReviewResponseFormat(suggestions)
            : undefined,
          temperature: 0,
          maxTokens: 400,
        });

        if (!content.trim()) {
          lastReason = "empty response";
        } else {
          const parsed = extractJsonObjectStrict(content);
          if (!parsed) {
            lastReason = "JSON parse failed";
          } else {
            const checked = validateReviewPayload(parsed, suggestions);
            if (checked.ok) {
              return mergeSuggestionsWithReviews(suggestions, checked.value);
            }
            lastReason = checked.reason;
          }
        }
      } catch (err) {
        lastReason = errorToMessage(err);
        if (useStructuredOutput && looksLikeStructuredOutputUnsupported(err)) {
          useStructuredOutput = false;
        }
      }

      if (attempt < RETRY_MAX) {
        await sleep(RETRY_DELAY_MS * attempt);
      }
    }

    // fallback: if secondary review fails, keep suggestions as manual-review grade
    console.warn("[local-lm-tag-generator] secondary review fallback:", lastReason);
    return suggestions.map((s) => ({
      ...s,
      granularityScore: 60,
      toneScore: 60,
      verdict: "review",
      reviewReason: `Secondary review failed: ${lastReason}`,
      selected: false,
    }));
  }

  async applyTagsToFrontmatter(
    file: TFile,
    tags: string[],
  ): Promise<{ added: string[]; merged: string[] }> {
    const normalizedIncoming = uniqueTags(tags);

    let added: string[] = [];
    let merged: string[] = [];

    await this.app.fileManager.processFrontMatter(file, (fm) => {
      const existing = normalizeTagsFromFrontmatter(fm.tags);
      merged = uniqueTags([...existing, ...normalizedIncoming]);

      const existingSet = new Set(existing.map((t) => t.toLowerCase()));
      added = merged.filter((t) => !existingSet.has(t.toLowerCase()));

      fm.tags = merged;
    });

    return { added, merged };
  }

  refreshSidebarViews(): void {
    const leaves = this.app.workspace.getLeavesOfType(VIEW_TYPE_TAG_GENERATOR);
    for (const leaf of leaves) {
      if (leaf.view instanceof TagGeneratorSidebarView) {
        void leaf.view.render();
      }
    }
  }

  async loadSettings(): Promise<void> {
    const loaded = (await this.loadData()) as Partial<PluginSettings> | null;
    const merged: PluginSettings = {
      ...DEFAULT_SETTINGS,
      ...(loaded ?? {}),
    };

    merged.newSuggestionCount = clamp(
      Number(merged.newSuggestionCount),
      MIN_NEW_SUGGESTIONS,
      MAX_NEW_SUGGESTIONS,
    );

    if (!isNewTagMode(merged.newTagMode)) {
      merged.newTagMode = DEFAULT_SETTINGS.newTagMode;
    }

    this.settings = merged;
  }

  async saveSettings(): Promise<void> {
    await this.saveData(this.settings);
  }
}

type GenerationMessageArgs = {
  file: TFile;
  noteText: string;
  allowedTags: string[];
  mode: NewTagMode;
  suggestionCount: number;
  maxExistingTags: number;
  attempt: number;
};

function buildGenerationMessages(args: GenerationMessageArgs): ChatMessage[] {
  const {
    file,
    noteText,
    allowedTags,
    mode,
    suggestionCount,
    maxExistingTags,
    attempt,
  } = args;

  const retryHint =
    attempt > 1
      ? "Previous output was invalid. Return strictly valid JSON now."
      : "";

  const modeRule =
    mode === "none"
      ? "new_tag_suggestions must be an empty array always."
      : [
          "If existing tags are enough, set coverage='sufficient' and return new_tag_suggestions as [].",
          `If coverage is insufficient, set coverage='insufficient' and return ${MIN_NEW_SUGGESTIONS}-${MAX_NEW_SUGGESTIONS} new tag suggestions (target about ${suggestionCount}).`,
        ].join("\n");

  const systemPrompt = [
    "You are a tag planner for Obsidian notes.",
    "Pick existing_tags only from the allowed list.",
    "Do not invent tags in existing_tags.",
    `existing_tags max: ${maxExistingTags}`,
    "Output JSON object only.",
    'Required shape: {"existing_tags":[],"coverage":"sufficient|insufficient","new_tag_suggestions":[{"tag":"...","reason":"...","confidence":0.0}]}.',
    "For new_tag_suggestions, keep wording style and granularity close to allowed tags.",
    "No markdown. No prose.",
    modeRule,
    retryHint,
  ]
    .filter(Boolean)
    .join("\n");

  const userPrompt = [
    `File name: ${file.basename}`,
    "",
    "Allowed tags:",
    allowedTags.map((t) => `- ${t}`).join("\n"),
    "",
    "Note body:",
    noteText,
  ].join("\n");

  return [
    { role: "system", content: systemPrompt },
    { role: "user", content: userPrompt },
  ];
}

type ReviewMessageArgs = {
  file: TFile;
  noteText: string;
  allowedTags: string[];
  existingTags: string[];
  suggestions: NewTagSuggestion[];
  attempt: number;
};

function buildReviewMessages(args: ReviewMessageArgs): ChatMessage[] {
  const { file, noteText, allowedTags, existingTags, suggestions, attempt } = args;

  const retryHint =
    attempt > 1
      ? "Previous output was invalid. Return strictly valid JSON now."
      : "";

  const systemPrompt = [
    "You are a secondary reviewer for new tag suggestions.",
    "Evaluate each candidate tag against allowed tags.",
    "Focus on granularity and tone similarity.",
    "Use verdict rules:",
    "- accept: both granularity and tone are close",
    "- review: borderline but potentially useful",
    "- reject: too far in granularity or tone",
    'Output JSON object only with this shape: {"reviews":[{"tag":"...","granularity_score":0-100,"tone_score":0-100,"verdict":"accept|review|reject","reason":"..."}]}.',
    "Return one review entry per candidate tag.",
    "No markdown. No prose.",
    retryHint,
  ]
    .filter(Boolean)
    .join("\n");

  const userPrompt = [
    `File name: ${file.basename}`,
    "",
    "Allowed tags:",
    allowedTags.map((t) => `- ${t}`).join("\n"),
    "",
    "Selected existing tags:",
    existingTags.map((t) => `- ${t}`).join("\n"),
    "",
    "Candidate new tags:",
    suggestions
      .map((s) => `- ${s.tag} (confidence=${Math.round(s.confidence * 100)}, reason=${s.reason})`)
      .join("\n"),
    "",
    "Note body:",
    noteText,
  ].join("\n");

  return [
    { role: "system", content: systemPrompt },
    { role: "user", content: userPrompt },
  ];
}

function buildGenerationResponseFormat(allowedTags: string[]): Record<string, unknown> {
  return {
    type: "json_schema",
    json_schema: {
      name: "tag_generation_plan",
      schema: {
        type: "object",
        additionalProperties: false,
        properties: {
          existing_tags: {
            type: "array",
            items: {
              type: "string",
              enum: allowedTags,
            },
            minItems: 0,
            maxItems: MAX_GENERATED_TAGS,
          },
          coverage: {
            type: "string",
            enum: ["sufficient", "insufficient"],
          },
          new_tag_suggestions: {
            type: "array",
            minItems: 0,
            maxItems: MAX_NEW_SUGGESTIONS,
            items: {
              type: "object",
              additionalProperties: false,
              properties: {
                tag: { type: "string" },
                reason: { type: "string" },
                confidence: {
                  type: "number",
                  minimum: 0,
                  maximum: 1,
                },
              },
              required: ["tag", "reason", "confidence"],
            },
          },
        },
        required: ["existing_tags", "coverage", "new_tag_suggestions"],
      },
    },
  };
}

function buildReviewResponseFormat(
  suggestions: NewTagSuggestion[],
): Record<string, unknown> {
  const tagEnum = suggestions.map((s) => s.tag);

  return {
    type: "json_schema",
    json_schema: {
      name: "tag_review_result",
      schema: {
        type: "object",
        additionalProperties: false,
        properties: {
          reviews: {
            type: "array",
            minItems: 0,
            maxItems: tagEnum.length,
            items: {
              type: "object",
              additionalProperties: false,
              properties: {
                tag: { type: "string", enum: tagEnum },
                granularity_score: {
                  type: "number",
                  minimum: 0,
                  maximum: 100,
                },
                tone_score: {
                  type: "number",
                  minimum: 0,
                  maximum: 100,
                },
                verdict: {
                  type: "string",
                  enum: ["accept", "review", "reject"],
                },
                reason: { type: "string" },
              },
              required: [
                "tag",
                "granularity_score",
                "tone_score",
                "verdict",
                "reason",
              ],
            },
          },
        },
        required: ["reviews"],
      },
    },
  };
}

async function callLmStudioChat(args: {
  messages: ChatMessage[];
  responseFormat?: Record<string, unknown>;
  temperature: number;
  maxTokens: number;
}): Promise<string> {
  const normalizedBase = LMSTUDIO_BASE_URL.replace(/\/$/, "");
  const baseWithV1 = /\/v1$/i.test(normalizedBase)
    ? normalizedBase
    : `${normalizedBase}/v1`;
  const url = `${baseWithV1}/chat/completions`;

  const payload: Record<string, unknown> = {
    model: LMSTUDIO_MODEL,
    temperature: args.temperature,
    max_tokens: args.maxTokens,
    messages: args.messages,
  };

  if (args.responseFormat) {
    payload.response_format = args.responseFormat;
  }

  const res = await requestUrl({
    url,
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: "Bearer lm-studio",
    },
    body: JSON.stringify(payload),
  });

  if (res.status >= 400) {
    throw new LmStudioRequestError(
      `LM Studio HTTP ${res.status}`,
      res.status,
      String(res.text ?? ""),
    );
  }

  let data: unknown;
  try {
    data = res.json ?? JSON.parse(res.text);
  } catch (err) {
    throw new LmStudioRequestError(
      `Unable to parse LM Studio JSON response: ${errorToMessage(err)}`,
      res.status,
      String(res.text ?? ""),
    );
  }

  if (
    typeof data === "object" &&
    data !== null &&
    "choices" in data &&
    Array.isArray((data as { choices?: unknown }).choices)
  ) {
    const firstChoice = (
      data as { choices: Array<{ message?: { content?: unknown } }> }
    ).choices[0];

    const content = firstChoice?.message?.content;
    if (typeof content === "string") return content;
  }

  return "";
}

function validateGenerationPayload(
  raw: Record<string, unknown>,
  allowedTags: string[],
  mode: NewTagMode,
  desiredSuggestionCount: number,
): ValidationResult<TagGenerationPayload> {
  const allowedMap = buildAllowedTagMap(allowedTags);

  const existingRaw = raw.existing_tags;
  if (!Array.isArray(existingRaw)) {
    return { ok: false, reason: "existing_tags is not an array" };
  }

  const existingTags: string[] = [];
  const seenExisting = new Set<string>();

  for (const item of existingRaw) {
    if (typeof item !== "string") {
      return { ok: false, reason: "existing_tags has non-string value" };
    }

    const normalized = normalizeTag(item);
    if (!normalized) continue;

    const canonical = allowedMap.get(normalized.toLowerCase());
    if (!canonical) {
      return { ok: false, reason: `existing tag not allowed: ${item}` };
    }

    const key = canonical.toLowerCase();
    if (seenExisting.has(key)) continue;
    seenExisting.add(key);
    existingTags.push(canonical);
  }

  if (existingTags.length > MAX_GENERATED_TAGS) {
    return {
      ok: false,
      reason: `existing_tags exceeds limit (${MAX_GENERATED_TAGS})`,
    };
  }

  const coverageRaw = raw.coverage;
  if (coverageRaw !== "sufficient" && coverageRaw !== "insufficient") {
    return { ok: false, reason: "coverage is invalid" };
  }
  const coverage: CoverageState = coverageRaw;

  const suggestionsRaw = raw.new_tag_suggestions;
  if (!Array.isArray(suggestionsRaw)) {
    return { ok: false, reason: "new_tag_suggestions is not an array" };
  }

  const existingSetLower = new Set(existingTags.map((t) => t.toLowerCase()));
  const allowedSetLower = new Set(allowedTags.map((t) => normalizeTag(t).toLowerCase()));

  const suggestions: NewTagSuggestion[] = [];
  const seenNew = new Set<string>();

  for (const item of suggestionsRaw) {
    if (!item || typeof item !== "object" || Array.isArray(item)) {
      return { ok: false, reason: "new_tag_suggestions has invalid item" };
    }

    const record = item as Record<string, unknown>;
    const tagRaw = record.tag;
    const reasonRaw = record.reason;
    const confidenceRaw = record.confidence;

    if (typeof tagRaw !== "string") {
      return { ok: false, reason: "new suggestion tag is not string" };
    }

    const normalizedTag = normalizeTag(tagRaw);
    if (!normalizedTag) continue;

    const tagKey = normalizedTag.toLowerCase();
    if (allowedSetLower.has(tagKey)) continue;
    if (existingSetLower.has(tagKey)) continue;
    if (seenNew.has(tagKey)) continue;

    const reason =
      typeof reasonRaw === "string" && reasonRaw.trim().length > 0
        ? reasonRaw.trim()
        : "";

    const confidence = parseConfidence(confidenceRaw);

    suggestions.push({
      tag: normalizedTag,
      reason,
      confidence,
    });
    seenNew.add(tagKey);
  }

  if (mode === "none") {
    return {
      ok: true,
      value: {
        existingTags,
        coverage,
        newSuggestions: [],
      },
    };
  }

  if (coverage === "sufficient") {
    return {
      ok: true,
      value: {
        existingTags,
        coverage,
        newSuggestions: [],
      },
    };
  }

  // coverage is insufficient and new tags are enabled
  if (
    suggestions.length < MIN_NEW_SUGGESTIONS ||
    suggestions.length > MAX_NEW_SUGGESTIONS
  ) {
    return {
      ok: false,
      reason: `new suggestions count must be ${MIN_NEW_SUGGESTIONS}-${MAX_NEW_SUGGESTIONS} when coverage is insufficient`,
    };
  }

  // bias towards desired count but keep strict range
  suggestions.sort((a, b) => b.confidence - a.confidence);
  const sliced = suggestions.slice(0, desiredSuggestionCount);

  if (sliced.length < MIN_NEW_SUGGESTIONS) {
    return {
      ok: false,
      reason: "not enough high-quality new suggestions",
    };
  }

  return {
    ok: true,
    value: {
      existingTags,
      coverage,
      newSuggestions: sliced,
    },
  };
}

interface ReviewRecord {
  tag: string;
  granularityScore: number;
  toneScore: number;
  verdict: ReviewVerdict;
  reason: string;
}

function validateReviewPayload(
  raw: Record<string, unknown>,
  suggestions: NewTagSuggestion[],
): ValidationResult<ReviewRecord[]> {
  const reviewsRaw = raw.reviews;
  if (!Array.isArray(reviewsRaw)) {
    return { ok: false, reason: "reviews is not an array" };
  }

  const suggestionMap = new Map<string, NewTagSuggestion>();
  for (const s of suggestions) {
    suggestionMap.set(s.tag.toLowerCase(), s);
  }

  const out = new Map<string, ReviewRecord>();

  for (const item of reviewsRaw) {
    if (!item || typeof item !== "object" || Array.isArray(item)) {
      continue;
    }

    const record = item as Record<string, unknown>;
    const tagRaw = record.tag;
    const granularityRaw = record.granularity_score;
    const toneRaw = record.tone_score;
    const verdictRaw = record.verdict;
    const reasonRaw = record.reason;

    if (typeof tagRaw !== "string") continue;

    const normalizedTag = normalizeTag(tagRaw);
    const key = normalizedTag.toLowerCase();
    if (!suggestionMap.has(key)) continue;

    const granularityScore = clampNumber(granularityRaw, 0, 100, 50);
    const toneScore = clampNumber(toneRaw, 0, 100, 50);
    const verdict = parseVerdict(verdictRaw, granularityScore, toneScore);
    const reason =
      typeof reasonRaw === "string" && reasonRaw.trim().length > 0
        ? reasonRaw.trim()
        : "No reason";

    out.set(key, {
      tag: normalizedTag,
      granularityScore,
      toneScore,
      verdict,
      reason,
    });
  }

  if (out.size === 0) {
    return { ok: false, reason: "review output had no usable rows" };
  }

  return { ok: true, value: Array.from(out.values()) };
}

function mergeSuggestionsWithReviews(
  suggestions: NewTagSuggestion[],
  reviews: ReviewRecord[],
): ReviewedSuggestion[] {
  const reviewMap = new Map<string, ReviewRecord>();
  for (const review of reviews) {
    reviewMap.set(review.tag.toLowerCase(), review);
  }

  const out: ReviewedSuggestion[] = [];

  for (const suggestion of suggestions) {
    const key = suggestion.tag.toLowerCase();
    const review = reviewMap.get(key);

    if (review) {
      out.push({
        ...suggestion,
        granularityScore: review.granularityScore,
        toneScore: review.toneScore,
        verdict: review.verdict,
        reviewReason: review.reason,
        selected: false,
      });
      continue;
    }

    out.push({
      ...suggestion,
      granularityScore: 60,
      toneScore: 60,
      verdict: "review",
      reviewReason: "No secondary review output for this tag",
      selected: false,
    });
  }

  return out;
}

function parseAllowedTagFile(raw: string): string[] {
  const lines = raw.split(/\r?\n/);
  const out: string[] = [];

  for (let line of lines) {
    line = line.trim();
    if (!line) continue;

    if (/^#{1,6}\s+/.test(line)) continue;

    line = line.replace(/^[-*+]\s+/, "");
    line = line.replace(/^\[[ xX]\]\s+/, "");
    line = line.replace(/^tags?\s*:\s*/i, "");
    if (!line) continue;

    const parts = line.split(",").map((s) => s.trim()).filter(Boolean);
    for (const part of parts) {
      const normalized = normalizeTag(part);
      if (normalized) out.push(normalized);
    }
  }

  return uniqueTags(out);
}

function renderAllowedTagFile(tags: string[]): string {
  const normalized = uniqueTags(tags);
  const body = normalized.map((t) => `#${t}`).join("\n");
  return `# Allowed tags\n\n${body}\n`;
}

function normalizeTag(raw: string): string {
  let tag = String(raw ?? "").trim();
  if (!tag) return "";

  tag = tag.replace(/^#+/, "").trim();
  tag = tag.replace(/^["'`]+|["'`]+$/g, "").trim();
  tag = tag.replace(/[\[\],]/g, "").trim();
  tag = tag.replace(/\s+/g, "-");
  tag = tag.replace(/[。、；;:]+$/g, "").trim();

  return tag;
}

function uniqueTags(tags: string[]): string[] {
  const out: string[] = [];
  const seen = new Set<string>();

  for (const tag of tags) {
    const normalized = normalizeTag(tag);
    if (!normalized) continue;

    const key = normalized.toLowerCase();
    if (seen.has(key)) continue;

    seen.add(key);
    out.push(normalized);
  }

  return out;
}

function normalizeTagsFromFrontmatter(value: unknown): string[] {
  if (Array.isArray(value)) {
    return uniqueTags(
      value
        .map((item) => (typeof item === "string" ? item : String(item)))
        .map(normalizeTag)
        .filter(Boolean),
    );
  }

  if (typeof value === "string") {
    const parts = value.split(",").map((s) => s.trim()).filter(Boolean);
    return uniqueTags(parts.map(normalizeTag).filter(Boolean));
  }

  return [];
}

function stripYamlFrontmatter(text: string): string {
  const raw = String(text ?? "");
  return raw.replace(/^---\r?\n[\s\S]*?\r?\n---\r?\n?/, "");
}

function trimText(text: string, maxChars: number): string {
  if (text.length <= maxChars) return text;
  return text.slice(0, maxChars) + "\n\n[TRUNCATED]";
}

function extractJsonObjectStrict(text: string): Record<string, unknown> | null {
  const raw = String(text ?? "").trim();
  if (!raw) return null;

  try {
    const parsed = JSON.parse(raw);
    if (parsed && typeof parsed === "object" && !Array.isArray(parsed)) {
      return parsed as Record<string, unknown>;
    }
  } catch {
    // Continue to fallback parsers.
  }

  const codeBlock = raw.match(/```(?:json)?\s*([\s\S]*?)\s*```/i);
  if (codeBlock?.[1]) {
    try {
      const parsed = JSON.parse(codeBlock[1]);
      if (parsed && typeof parsed === "object" && !Array.isArray(parsed)) {
        return parsed as Record<string, unknown>;
      }
    } catch {
      // Continue.
    }
  }

  const objectMatch = raw.match(/\{[\s\S]*\}/);
  if (objectMatch?.[0]) {
    try {
      const parsed = JSON.parse(objectMatch[0]);
      if (parsed && typeof parsed === "object" && !Array.isArray(parsed)) {
        return parsed as Record<string, unknown>;
      }
    } catch {
      // No valid object found.
    }
  }

  return null;
}

function parseConfidence(raw: unknown): number {
  if (typeof raw === "number" && Number.isFinite(raw)) {
    return clamp(raw, 0, 1);
  }

  if (typeof raw === "string") {
    const parsed = Number(raw);
    if (Number.isFinite(parsed)) return clamp(parsed, 0, 1);
  }

  return 0.5;
}

function parseVerdict(
  raw: unknown,
  granularityScore: number,
  toneScore: number,
): ReviewVerdict {
  if (raw === "accept" || raw === "review" || raw === "reject") {
    return raw;
  }

  if (granularityScore >= 75 && toneScore >= 75) return "accept";
  if (granularityScore < 45 || toneScore < 45) return "reject";
  return "review";
}

function scoreSuggestion(suggestion: ReviewedSuggestion): number {
  const verdictBase =
    suggestion.verdict === "accept"
      ? 200
      : suggestion.verdict === "review"
        ? 100
        : 0;

  return (
    verdictBase +
    suggestion.granularityScore +
    suggestion.toneScore +
    Math.round(suggestion.confidence * 100)
  );
}

function buildAllowedTagMap(allowedTags: string[]): Map<string, string> {
  const out = new Map<string, string>();
  for (const tag of allowedTags) {
    const normalized = normalizeTag(tag);
    if (!normalized) continue;
    const key = normalized.toLowerCase();
    if (!out.has(key)) out.set(key, normalized);
  }
  return out;
}

function verdictLabel(verdict: ReviewVerdict): string {
  if (verdict === "accept") return "適合";
  if (verdict === "review") return "要確認";
  return "不採用";
}

function looksLikeStructuredOutputUnsupported(err: unknown): boolean {
  const message = errorToMessage(err).toLowerCase();
  return (
    message.includes("response_format") ||
    message.includes("json_schema") ||
    message.includes("structured") ||
    message.includes("unsupported")
  );
}

function errorToMessage(err: unknown): string {
  if (err instanceof LmStudioRequestError) {
    const body = (err.bodyText ?? "").slice(0, 400);
    return body ? `${err.message}: ${body}` : err.message;
  }
  if (err instanceof Error) return err.message;
  return String(err);
}

function isNewTagMode(value: unknown): value is NewTagMode {
  return value === "none" || value === "approval" || value === "auto";
}

function clamp(value: number, min: number, max: number): number {
  if (!Number.isFinite(value)) return min;
  return Math.max(min, Math.min(max, value));
}

function clampNumber(raw: unknown, min: number, max: number, fallback: number): number {
  if (typeof raw === "number" && Number.isFinite(raw)) {
    return clamp(raw, min, max);
  }

  if (typeof raw === "string") {
    const parsed = Number(raw);
    if (Number.isFinite(parsed)) return clamp(parsed, min, max);
  }

  return fallback;
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
