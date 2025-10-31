/* global Office, OfficeRuntime, Word */
/**
 * OpenAI Word Assistant Task Pane Script
 * Implements requirements: Word integration, OpenAI fetch, storage, telemetry.
 */

(() => {
  /** @type {HTMLSelectElement} */
  let operationSelect;
  /** @type {HTMLSelectElement} */
  let toneSelect;
  /** @type {HTMLInputElement} */
  let languageInput;
  /** @type {HTMLTextAreaElement} */
  let promptInput;
  /** @type {HTMLInputElement} */
  let modelInput;
  /** @type {HTMLInputElement} */
  let apiKeyInput;
  /** @type {HTMLButtonElement} */
  let saveKeyButton;
  /** @type {HTMLButtonElement} */
  let clearKeyButton;
  /** @type {HTMLButtonElement} */
  let runSelectionButton;
  /** @type {HTMLButtonElement} */
  let replaceSelectionButton;
  /** @type {HTMLButtonElement} */
  let insertCursorButton;
  /** @type {HTMLDivElement} */
  let statusDiv;
  /** @type {HTMLUListElement} */
  let telemetryList;

  const STORAGE_KEY = "openai-word-assistant-api-key";
  const MAX_TELEMETRY = 5;

  /**
   * Human-readable operation descriptors for prompt construction (requirement: helper builds prompt).
   * @type {Record<string, string>}
   */
  const OPERATION_DESCRIPTIONS = {
    rewrite: "Rewrite the supplied content to be clearer and more concise.",
    summarize: "Summarize the supplied content in a few sentences.",
    translate: "Translate the supplied content.",
    "fix-grammar": "Fix grammar, spelling, and punctuation issues in the supplied content.",
  };

  /** Optional tone guidance appended to the prompt (nice-to-have). */
  const TONE_DESCRIPTIONS = {
    formal: "Use a formal tone.",
    friendly: "Use a friendly, approachable tone.",
    technical: "Use precise, technical language appropriate for expert readers.",
  };

  /**
   * Module encapsulating OpenAI configuration (requirement: single module for base URL + endpoint).
   */
  const OpenAIClient = (() => {
    const config = {
      baseUrl: "https://api.openai.com/v1",
      endpoint: "/chat/completions",
    };

    /**
     * Call OpenAI Chat Completions API.
     * @param {string} apiKey
     * @param {string} model
     * @param {Array<{role: string, content: string}>} messages
     * @returns {Promise<{text: string, usage?: any}>}
     */
    async function call(apiKey, model, messages) {
      const url = `${config.baseUrl}${config.endpoint}`;
      const response = await fetch(url, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${apiKey}`,
        },
        body: JSON.stringify({
          model,
          messages,
          temperature: 0.4,
        }),
      });

      if (!response.ok) {
        let details = "";
        try {
          const data = await response.json();
          if (data?.error?.message) {
            details = data.error.message;
          }
        } catch (error) {
          // ignore parse failures
        }

        const friendly = buildFriendlyError(response.status, details);
        throw new Error(friendly);
      }

      const payload = await response.json();
      const text = payload?.choices?.[0]?.message?.content?.trim();
      if (!text) {
        throw new Error("OpenAI returned an empty response. Try again or adjust your prompt.");
      }
      return { text, usage: payload?.usage };
    }

    return { config, call };
  })();

  /**
   * Requirement: Show clear error messages for common HTTP codes.
   * @param {number} status
   * @param {string} details
   */
  function buildFriendlyError(status, details) {
    const suffix = details ? ` Details: ${details}` : "";
    switch (status) {
      case 401:
        return `Authentication failed (401). Check that your API key is valid and has billing enabled.${suffix}`;
      case 403:
        return `Access forbidden (403). Ensure the model '${modelInput ? modelInput.value : "specified"}' is available to your account.${suffix}`;
      case 404:
        return `Endpoint or model not found (404). Verify the model name and endpoint.${suffix}`;
      case 429:
        return `Rate limit exceeded (429). Wait a moment before retrying or reduce request frequency.${suffix}`;
      case 500:
      case 503:
        return `OpenAI service is unavailable (${status}). Try again later.${suffix}`;
      default:
        return `Request failed with status ${status}.${suffix}`;
    }
  }

  /**
   * Requirement: Store API key securely using OfficeRuntime.storage with localStorage fallback.
   * @returns {Promise<void>}
   */
  async function saveApiKey(key) {
    if (!key) {
      throw new Error("API key is empty");
    }
    if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage) {
      await OfficeRuntime.storage.setItem(STORAGE_KEY, key);
    } else {
      // LocalStorage fallback with security reminder already present in UI copy.
      localStorage.setItem(STORAGE_KEY, key);
    }
  }

  /** Load saved API key. */
  async function loadApiKey() {
    if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage) {
      const value = await OfficeRuntime.storage.getItem(STORAGE_KEY);
      return value || "";
    }
    return localStorage.getItem(STORAGE_KEY) || "";
  }

  /** Clear stored key. */
  async function clearApiKey() {
    if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage) {
      await OfficeRuntime.storage.removeItem(STORAGE_KEY);
    } else {
      localStorage.removeItem(STORAGE_KEY);
    }
  }

  /** Set status text and usage metadata (requirement: status area). */
  function setStatus(message, options) {
    const { isError = false, usage } = options || {};
    if (!statusDiv) {
      return;
    }
    const lines = [];
    if (message) {
      lines.push(isError ? `⚠️ ${message}` : message);
    }
    if (usage) {
      const parts = [];
      if (typeof usage.prompt_tokens === "number") {
        parts.push(`Prompt tokens: ${usage.prompt_tokens}`);
      }
      if (typeof usage.completion_tokens === "number") {
        parts.push(`Completion tokens: ${usage.completion_tokens}`);
      }
      if (typeof usage.total_tokens === "number") {
        parts.push(`Total tokens: ${usage.total_tokens}`);
      }
      if (parts.length > 0) {
        lines.push(`Usage → ${parts.join(", ")}`);
      }
    }
    statusDiv.textContent = lines.join("\n");
    statusDiv.classList.toggle("error", isError);
  }

  /** Adds entry to telemetry panel (optional nice-to-have). */
  const telemetry = [];
  function pushTelemetry(entry) {
    telemetry.unshift(entry);
    while (telemetry.length > MAX_TELEMETRY) {
      telemetry.pop();
    }
    renderTelemetry();
  }

  function renderTelemetry() {
    if (!telemetryList) return;
    telemetryList.innerHTML = "";
    telemetry.forEach((item) => {
      const li = document.createElement("li");
      li.textContent = `${item.timestamp} · ${item.operation} → ${item.outcome}`;
      telemetryList.appendChild(li);
    });
  }

  /**
   * Builds the user content string for OpenAI (requirement: helper building prompt).
   * @param {string} operation
   * @param {string} prompt
   * @param {string} selectedText
   * @param {string} tone
   * @param {string} language
   */
  function buildUserContent(operation, prompt, selectedText, tone, language) {
    const operationDescription = OPERATION_DESCRIPTIONS[operation] || operation;
    const segments = [operationDescription];
    if (prompt) {
      segments.push(`Instruction: ${prompt}`);
    }
    if (tone && TONE_DESCRIPTIONS[tone]) {
      segments.push(TONE_DESCRIPTIONS[tone]);
    }
    if (language) {
      segments.push(`Respond in the ${language} language.`);
    }
    if (selectedText) {
      segments.push("Selected text:\n" + selectedText);
    } else {
      segments.push("No selection was provided. Work with the user instruction only.");
    }
    return segments.join("\n\n");
  }

  /** Builds the request payload for OpenAI. */
  function buildMessages(userContent) {
    return [
      {
        role: "system",
        content: "You are a helpful writing assistant operating inside Microsoft Word.",
      },
      {
        role: "user",
        content: userContent,
      },
    ];
  }

  /**
   * Reads current selection text (requirement: Word integration).
   * @returns {Promise<string>}
   */
  async function getSelectionText() {
    return Word.run(async (context) => {
      const range = context.document.getSelection();
      range.load("text");
      await context.sync();
      return range.text || "";
    });
  }

  /**
   * Inserts text relative to the selection based on mode.
   * @param {"replace"|"after"|"cursor"} mode
   * @param {string} text
   */
  async function insertText(mode, text) {
    return Word.run(async (context) => {
      const range = context.document.getSelection();
      if (mode === "replace") {
        range.insertText(text, Word.InsertLocation.replace);
      } else if (mode === "after") {
        range.insertParagraph(text, Word.InsertLocation.after);
      } else {
        range.insertText(text, Word.InsertLocation.replace);
      }
      await context.sync();
    });
  }

  /** Retrieves stored API key or shows error. */
  async function requireApiKey() {
    const key = apiKeyInput.value.trim();
    if (!key) {
      throw new Error("Add your OpenAI API key in the API Settings section before running an operation.");
    }
    return key;
  }

  /**
   * Handles the OpenAI call for the provided Word selection context.
   * @param {"replace"|"after"|"cursor"} mode
   */
  async function handleRequest(mode) {
    try {
      disableActionButtons(true);
      setStatus("Calling OpenAI...", { isError: false });

      const selectionText = mode === "cursor" ? "" : (await getSelectionText());

      if (!selectionText && mode !== "cursor") {
        setStatus(
          "The current selection is empty. Use Insert at Cursor or select some text before running.",
          { isError: true }
        );
        disableActionButtons(false);
        return;
      }

      const apiKey = await requireApiKey();
      const operation = operationSelect.value;
      const prompt = promptInput.value.trim();
      const tone = toneSelect.value;
      const language = languageInput.value.trim();
      const model = modelInput.value.trim();

      const userContent = buildUserContent(operation, prompt, selectionText, tone, language);
      const messages = buildMessages(userContent);

      const startedAt = new Date();
      const { text, usage } = await OpenAIClient.call(apiKey, model, messages);
      await insertText(mode === "cursor" ? "cursor" : mode === "replace" ? "replace" : "after", text);

      const finishedAt = new Date();
      const duration = (finishedAt.getTime() - startedAt.getTime()) / 1000;
      setStatus(`Success in ${duration.toFixed(1)}s`, { usage });
      pushTelemetry({
        timestamp: finishedAt.toLocaleTimeString(),
        operation: `${operation} (${mode})`,
        outcome: `${text.slice(0, 40)}${text.length > 40 ? "…" : ""}`,
      });
    } catch (error) {
      console.error(error);
      setStatus(error.message || String(error), { isError: true });
      pushTelemetry({
        timestamp: new Date().toLocaleTimeString(),
        operation: `${operationSelect.value} (${mode})`,
        outcome: "Error",
      });
    } finally {
      disableActionButtons(false);
    }
  }

  /** Disable/enable action buttons during async operations (requirement: graceful async handling). */
  function disableActionButtons(disabled) {
    runSelectionButton.disabled = disabled;
    replaceSelectionButton.disabled = disabled;
    insertCursorButton.disabled = disabled;
  }

  /** Initialize UI, load stored key, and wire handlers. */
  async function initialize() {
    operationSelect = document.getElementById("operation");
    toneSelect = document.getElementById("tone");
    languageInput = document.getElementById("language");
    promptInput = document.getElementById("prompt");
    modelInput = document.getElementById("model");
    apiKeyInput = document.getElementById("api-key");
    saveKeyButton = document.getElementById("save-key");
    clearKeyButton = document.getElementById("clear-key");
    runSelectionButton = document.getElementById("run-selection");
    replaceSelectionButton = document.getElementById("replace-selection");
    insertCursorButton = document.getElementById("insert-cursor");
    statusDiv = document.getElementById("status");
    telemetryList = document.getElementById("telemetry");

    const savedKey = await loadApiKey();
    if (savedKey) {
      apiKeyInput.value = savedKey;
    }

    saveKeyButton.addEventListener("click", async () => {
      try {
        await saveApiKey(apiKeyInput.value.trim());
        setStatus("API key saved locally.");
      } catch (error) {
        setStatus(error.message || String(error), { isError: true });
      }
    });

    clearKeyButton.addEventListener("click", async () => {
      await clearApiKey();
      apiKeyInput.value = "";
      setStatus("API key cleared.");
    });

    runSelectionButton.addEventListener("click", () => handleRequest("after"));
    replaceSelectionButton.addEventListener("click", () => handleRequest("replace"));
    insertCursorButton.addEventListener("click", () => handleRequest("cursor"));

    setStatus("Ready.");
  }

  // Requirement: call initialize once Office is ready.
  Office.onReady(() => {
    if (document.readyState === "complete" || document.readyState === "interactive") {
      initialize().catch((error) => setStatus(error.message || String(error), { isError: true }));
    } else {
      document.addEventListener("DOMContentLoaded", () => {
        initialize().catch((error) => setStatus(error.message || String(error), { isError: true }));
      });
    }
  });
})();
