/* global Office, PowerPoint, fetch */

// ---------------------------------------------------------------------------
// Constants & element refs
// ---------------------------------------------------------------------------
const LS_KEY = "geminiSlides.settings";

const els = {};

function cacheElements() {
  [
    "settingsToggle", "settingsPanel",
    "apiKey", "toggleKeyVisibility",
    "modelName", "customModel", "saveSettings",
    "topic", "slideCount", "tone",
    "includeNotes",
    "generateBtn", "errorMsg", "progressList",
  ].forEach((id) => (els[id] = document.getElementById(id)));
}

// ---------------------------------------------------------------------------
// Settings (stored client-side only; never sent anywhere but Gemini's API)
// ---------------------------------------------------------------------------
function loadSettings() {
  try {
    const raw = localStorage.getItem(LS_KEY);
    return raw ? JSON.parse(raw) : {};
  } catch {
    return {};
  }
}

function saveSettings(settings) {
  localStorage.setItem(LS_KEY, JSON.stringify(settings));
}

function getSelectedModel() {
  return els.modelName.value === "custom"
    ? els.customModel.value.trim()
    : els.modelName.value;
}

function applySettingsToForm(settings) {
  if (settings.apiKey) els.apiKey.value = settings.apiKey;
  if (settings.model) {
    const known = Array.from(els.modelName.options).some((o) => o.value === settings.model);
    if (known) {
      els.modelName.value = settings.model;
      els.customModel.classList.add("hidden");
    } else {
      els.modelName.value = "custom";
      els.customModel.value = settings.model;
      els.customModel.classList.remove("hidden");
    }
  }
}

// ---------------------------------------------------------------------------
// UI helpers
// ---------------------------------------------------------------------------
function showError(message) {
  els.errorMsg.textContent = message;
  els.errorMsg.classList.remove("hidden");
}

function clearError() {
  els.errorMsg.classList.add("hidden");
  els.errorMsg.textContent = "";
}

function resetProgressList(slideTitles) {
  els.progressList.innerHTML = "";
  els.progressList.classList.remove("hidden");
  slideTitles.forEach((title, i) => {
    const li = document.createElement("li");
    li.id = `progress-${i}`;
    li.innerHTML = `<span class="dot"></span><span class="label"></span>`;
    li.querySelector(".label").textContent = title;
    els.progressList.appendChild(li);
  });
}

function markProgress(index, state) {
  const li = document.getElementById(`progress-${index}`);
  if (!li) return;
  li.classList.remove("active", "done", "error");
  li.classList.add(state);
}

function setBusy(isBusy) {
  els.generateBtn.disabled = isBusy;
  els.generateBtn.querySelector(".btn-label").textContent = isBusy
    ? "Generating…"
    : "Generate slides";
}

// ---------------------------------------------------------------------------
// Gemini call
// ---------------------------------------------------------------------------
// Asks Gemini for strict JSON describing the slide deck, using the API's
// structured-output mode (responseMimeType) so we get clean JSON back
// instead of markdown-fenced text.
async function generateDeckWithGemini({ apiKey, model, topic, slideCount, tone, includeNotes }) {
  const schemaHint = includeNotes
    ? `[{"title": string, "bullets": string[], "notes": string}, ...]`
    : `[{"title": string, "bullets": string[]}, ...]`;

  const prompt = `You are an expert presentation writer. Create the content for a
${slideCount}-slide PowerPoint deck about: "${topic}".

Tone: ${tone}.
Rules:
- Return ONLY a JSON array, no prose, matching this shape: ${schemaHint}
- Exactly ${slideCount} slide objects, in the order they should appear.
- Each slide's "title" is short (under 8 words).
- Each slide has 3-5 "bullets", each a single concise sentence or phrase (under 16 words), no bullet characters or numbering in the text itself.
${includeNotes ? '- Each slide has "notes": 1-3 sentences of speaker talking points expanding on the bullets.' : ""}
- The first slide should act as a title slide: put the deck title in "title" and put a one-line subtitle as the single entry in "bullets".`;

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(
    model
  )}:generateContent`;

  const response = await fetch(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-goog-api-key": apiKey,
    },
    body: JSON.stringify({
      contents: [{ role: "user", parts: [{ text: prompt }] }],
      generationConfig: {
        responseMimeType: "application/json",
      },
    }),
  });

  if (!response.ok) {
    const errBody = await response.text().catch(() => "");
    throw new Error(
      `Gemini API error (${response.status}). ${extractApiErrorMessage(errBody)}`
    );
  }

  const data = await response.json();
  const text = data?.candidates?.[0]?.content?.parts?.map((p) => p.text || "").join("") || "";

  if (!text) {
    const blockReason = data?.promptFeedback?.blockReason;
    throw new Error(
      blockReason
        ? `Gemini did not return content (blocked: ${blockReason}).`
        : "Gemini returned an empty response."
    );
  }

  return parseSlideJson(text);
}

function extractApiErrorMessage(rawBody) {
  try {
    const parsed = JSON.parse(rawBody);
    return parsed?.error?.message || rawBody.slice(0, 200);
  } catch {
    return rawBody.slice(0, 200) || "Check your API key and model name.";
  }
}

function parseSlideJson(text) {
  // Structured-output mode should return clean JSON, but strip any
  // ```json fences defensively in case a model ignores the config.
  const cleaned = text.trim().replace(/^```json\s*/i, "").replace(/```$/i, "").trim();
  let deck;
  try {
    deck = JSON.parse(cleaned);
  } catch (e) {
    throw new Error("Couldn't parse Gemini's response as JSON. Try again or switch models.");
  }
  if (!Array.isArray(deck) || deck.length === 0) {
    throw new Error("Gemini's response wasn't a non-empty array of slides.");
  }
  return deck.map((s) => ({
    title: String(s.title || "Untitled slide"),
    bullets: Array.isArray(s.bullets) ? s.bullets.map(String) : [],
    notes: s.notes ? String(s.notes) : "",
  }));
}

// ---------------------------------------------------------------------------
// PowerPoint insertion
// ---------------------------------------------------------------------------
// Layout constants for a standard 13.33" x 7.5" (widescreen) slide, in points.
const SLIDE = {
  titleTop: 40,
  titleHeight: 70,
  bodyTop: 130,
  bodyHeight: 330,
  notesTop: 470,
  notesHeight: 50,
  left: 50,
  width: 860,
};

async function buildDeckInPowerPoint(deck, { includeNotes }, onSlideDone) {
  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();

    // slides.add() always appends to the end.
    const startCount = slides.items.length;

    for (let i = 0; i < deck.length; i++) {
      const slideData = deck[i];

      // Add a new blank slide at the end of the deck, then get a fresh,
      // properly bound reference to it via getItemAt() — indexing into a
      // previously-loaded .items array (the old approach here) produces a
      // reference that PowerPoint on the web can fail to resolve later.
      slides.add();
      // eslint-disable-next-line no-await-in-loop
      await context.sync();

      const newIndex = startCount + i;
      const newSlide = slides.getItemAt(newIndex);
      const shapes = newSlide.shapes;

      const titleBox = shapes.addTextBox(slideData.title);
      titleBox.left = SLIDE.left;
      titleBox.top = SLIDE.titleTop;
      titleBox.width = SLIDE.width;
      titleBox.height = SLIDE.titleHeight;
      titleBox.name = "GeminiSlides_Title";
      titleBox.textFrame.textRange.font.size = 28;
      titleBox.textFrame.textRange.font.bold = true;
      // eslint-disable-next-line no-await-in-loop
      await context.sync();

      const bodyText = slideData.bullets.map((b) => `•  ${b}`).join("\n");
      if (bodyText) {
        const bodyBox = shapes.addTextBox(bodyText);
        bodyBox.left = SLIDE.left;
        bodyBox.top = SLIDE.bodyTop;
        bodyBox.width = SLIDE.width;
        bodyBox.height = SLIDE.bodyHeight;
        bodyBox.name = "GeminiSlides_Body";
        bodyBox.textFrame.textRange.font.size = 18;
        bodyBox.textFrame.wordWrap = true;
        // eslint-disable-next-line no-await-in-loop
        await context.sync();
      }

      // The PowerPoint JS API does not currently expose the real speaker-notes
      // pane, so as a practical stand-in we add a small on-slide note instead.
      if (includeNotes && slideData.notes) {
        const notesBox = shapes.addTextBox(`Notes: ${slideData.notes}`);
        notesBox.left = SLIDE.left;
        notesBox.top = SLIDE.notesTop;
        notesBox.width = SLIDE.width;
        notesBox.height = SLIDE.notesHeight;
        notesBox.name = "GeminiSlides_Notes";
        notesBox.textFrame.textRange.font.size = 10;
        notesBox.textFrame.textRange.font.italic = true;
        // eslint-disable-next-line no-await-in-loop
        await context.sync();
      }

      // eslint-disable-next-line no-await-in-loop
      await context.sync();
      onSlideDone(i);
    }

  });
}

// ---------------------------------------------------------------------------
// Main generate flow
// ---------------------------------------------------------------------------
async function handleGenerateClick() {
  clearError();

  const settings = loadSettings();
  const apiKey = els.apiKey.value.trim() || settings.apiKey;
  const model = getSelectedModel() || settings.model;
  const topic = els.topic.value.trim();
  const slideCount = Math.max(1, Math.min(20, parseInt(els.slideCount.value, 10) || 6));
  const tone = els.tone.value;
  const includeNotes = els.includeNotes.checked;

  if (!apiKey) {
    showError("Add your Gemini API key in Settings first.");
    els.settingsPanel.classList.remove("hidden");
    return;
  }
  if (!model) {
    showError("Choose or enter a Gemini model name in Settings.");
    els.settingsPanel.classList.remove("hidden");
    return;
  }
  if (!topic) {
    showError("Describe the topic or outline for your deck.");
    return;
  }

  // Persist settings used for this run.
  saveSettings({ apiKey, model });

  setBusy(true);
  resetProgressList(Array.from({ length: slideCount }, (_, i) => `Slide ${i + 1}`));

  try {
    const deck = await generateDeckWithGemini({
      apiKey,
      model,
      topic,
      slideCount,
      tone,
      includeNotes,
    });

    // Reconcile progress list with the actual number of slides returned.
    resetProgressList(deck.map((s, i) => s.title || `Slide ${i + 1}`));
    deck.forEach((_, i) => markProgress(i, "active"));

    await buildDeckInPowerPoint(deck, { includeNotes }, (i) => markProgress(i, "done"));
  } catch (err) {
    console.error(err);
    const detail = describeOfficeError(err);
    showError(detail);
    els.progressList.querySelectorAll("li:not(.done)").forEach((li) => li.classList.add("error"));
  } finally {
    setBusy(false);
  }
}

// Office.js errors (OfficeExtension.Error) carry a generic "GeneralException"
// name/message but often have far more detail in .code and .debugInfo — pull
// that out so the on-screen error is actually actionable.
function describeOfficeError(err) {
  const parts = [];
  if (err?.code) parts.push(`[${err.code}]`);
  parts.push(err?.message || "Something went wrong generating the deck.");
  if (err?.debugInfo?.errorLocation) {
    parts.push(`(at ${err.debugInfo.errorLocation})`);
  }
  console.error("Office error debugInfo:", err?.debugInfo);
  return parts.join(" ");
}

// ---------------------------------------------------------------------------
// Wiring
// ---------------------------------------------------------------------------
Office.onReady((info) => {
  if (info.host !== Office.HostType.PowerPoint) return;

  cacheElements();
  applySettingsToForm(loadSettings());

  els.settingsToggle.addEventListener("click", () => {
    els.settingsPanel.classList.toggle("hidden");
  });

  els.toggleKeyVisibility.addEventListener("click", () => {
    els.apiKey.type = els.apiKey.type === "password" ? "text" : "password";
  });

  els.modelName.addEventListener("change", () => {
    els.customModel.classList.toggle("hidden", els.modelName.value !== "custom");
  });

  els.saveSettings.addEventListener("click", () => {
    const apiKey = els.apiKey.value.trim();
    const model = getSelectedModel();
    if (!apiKey) {
      showError("Enter an API key before saving.");
      return;
    }
    saveSettings({ apiKey, model });
    clearError();
    els.settingsPanel.classList.add("hidden");
  });

  els.generateBtn.addEventListener("click", handleGenerateClick);
});
