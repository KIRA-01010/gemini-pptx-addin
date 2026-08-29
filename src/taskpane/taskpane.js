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
    "topic", "slideCount", "tone", "template", "accentColor",
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
        maxOutputTokens: 4096,
      },
    }),
  });

  if (!response.ok) {
    if (response.status === 429) {
      const errBody = await response.text().catch(() => "");
      const retryMatch = errBody.match(/retry in ([\d.]+)s/i);
      const waitMsg = retryMatch ? ` Try again in about ${Math.ceil(parseFloat(retryMatch[1]))}s.` : " Wait a bit and try again.";
      throw new Error(`You've hit Gemini's free-tier rate limit.${waitMsg}`);
    }
    const errBody = await response.text().catch(() => "");
    throw new Error(
      `Gemini API error (${response.status}). ${extractApiErrorMessage(errBody)}`
    );
  }

  const data = await response.json();
  const candidate = data?.candidates?.[0];
  const text = candidate?.content?.parts?.map((p) => p.text || "").join("") || "";

  if (!text) {
    const blockReason = data?.promptFeedback?.blockReason;
    const finishReason = candidate?.finishReason;
    throw new Error(
      blockReason
        ? `Gemini did not return content (blocked: ${blockReason}).`
        : finishReason
        ? `Gemini did not return content (finish reason: ${finishReason}).`
        : "Gemini returned an empty response."
    );
  }

  if (candidate?.finishReason === "MAX_TOKENS") {
    throw new Error(
      "Gemini's response was cut off before finishing (hit the token limit). Try fewer slides or shorter bullets."
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
    console.error("Raw Gemini response that failed to parse:", text);
    const snippet = cleaned.slice(0, 180).replace(/\s+/g, " ");
    throw new Error(
      `Couldn't parse Gemini's response as JSON: ${e.message}. Response started with: "${snippet}${cleaned.length > 180 ? "…" : ""}"`
    );
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
  width: 960,
  height: 540,
  accentBarHeight: 8,
  titleTop: 55,
  titleHeight: 65,
  ruleTop: 122,
  ruleHeight: 3,
  ruleWidth: 90,
  bodyTop: 145,
  bodyHeight: 320,
  notesTop: 478,
  notesHeight: 40,
  left: 55,
  contentWidth: 850,
  badgeSize: 30,
};

const COLORS = {
  dark: "#1C1B29",
  body: "#3F3D56",
  secondary: "#6E6B85",
  white: "#FFFFFF",
  darkBg: "#161522",
  darkBody: "#C9C7DA",
  cardBg: "#F4F3FA",
};

function flatFill(shape, color) {
  shape.fill.setSolidColor(color);
  shape.lineFormat.visible = false;
}

async function findBlankLayoutOptions() {
  let addOptions;
  await PowerPoint.run(async (context) => {
    const slideMasters = context.presentation.slideMasters;
    slideMasters.load("items/id");
    await context.sync();
    slideMasters.items.forEach((m) => m.layouts.load("items/id,items/name"));
    await context.sync();

    outer: for (const master of slideMasters.items) {
      for (const layout of master.layouts.items) {
        if (/blank/i.test(layout.name)) {
          addOptions = { slideMasterId: master.id, layoutId: layout.id };
          break outer;
        }
      }
    }
  });
  return addOptions;
}

async function addOneSlide(slideData, index, addOptions, { includeNotes, template, accentColor }) {
  const isTitleSlide = index === 0;

  // Step 1: create the slide and find its index — isolated on its own.
  let newIndex;
  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.add(addOptions);
    await context.sync();
    const countResult = slides.getCount();
    await context.sync();
    newIndex = countResult.value - 1;
  });

  // NOTE: PowerPoint.ShapeCollection has no addImage() method (confirmed —
  // it exists on Excel's ShapeCollection but not PowerPoint's), so real
  // images/gradients are not achievable here. Visual richness comes from
  // bold, layered geometric shapes instead — see drawDecorativeShapes().

  // Step 2: draw the rest of the slide's shapes in its own isolated batch.
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.slides.getItemAt(newIndex).shapes;

    const drawFn = TEMPLATES[template] || TEMPLATES.minimal;
    drawFn(shapes, slideData, {
      isTitleSlide,
      slideNumber: index + 1,
      accentColor,
      skipBackground: false,
    });

    // The PowerPoint JS API does not currently expose the real speaker-notes
    // pane, so as a practical stand-in we add a small on-slide note instead.
    if (includeNotes && slideData.notes) {
      const notesColor = template === "dark" ? COLORS.secondary : COLORS.secondary;
      const notesBox = shapes.addTextBox(`Notes: ${slideData.notes}`);
      notesBox.left = SLIDE.left;
      notesBox.top = SLIDE.notesTop;
      notesBox.width = SLIDE.contentWidth;
      notesBox.height = SLIDE.notesHeight;
      notesBox.name = "GeminiSlides_Notes";
      notesBox.textFrame.textRange.font.size = 10;
      notesBox.textFrame.textRange.font.italic = true;
      notesBox.textFrame.textRange.font.color = notesColor;
    }

    await context.sync();
  });
}

// ---------------------------------------------------------------------------
// Template layouts
// ---------------------------------------------------------------------------
// Each template function draws directly onto `shapes` (a slide's shape
// collection) and handles both the title-slide and content-slide case.

function drawBadge(shapes, slideNumber, { left, top, fillColor, textColor }) {
  const badge = shapes.addGeometricShape(PowerPoint.GeometricShapeType.oval);
  badge.left = left;
  badge.top = top;
  badge.width = SLIDE.badgeSize;
  badge.height = SLIDE.badgeSize;
  badge.name = "GeminiSlides_PageBadge";
  flatFill(badge, fillColor);
  badge.textFrame.textRange.text = String(slideNumber);
  badge.textFrame.textRange.font.size = 13;
  badge.textFrame.textRange.font.bold = true;
  badge.textFrame.textRange.font.color = textColor;
  badge.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middleCentered;
  badge.textFrame.textRange.paragraphFormat.horizontalAlignment =
    PowerPoint.ParagraphHorizontalAlignment.center;
}

function shadeColor(hex, percent) {
  // percent: -1 (black) to +1 (white); 0 = unchanged.
  const { r, g, b } = hexToRgb(hex);
  const mix = (channel) =>
    percent < 0
      ? Math.round(channel * (1 + percent))
      : Math.round(channel + (255 - channel) * percent);
  const clamp = (n) => Math.max(0, Math.min(255, n));
  const toHex = (n) => clamp(n).toString(16).padStart(2, "0");
  return `#${toHex(mix(r))}${toHex(mix(g))}${toHex(mix(b))}`;
}

function hexToRgb(hex) {
  const clean = hex.replace("#", "");
  const n = parseInt(clean, 16);
  return { r: (n >> 16) & 255, g: (n >> 8) & 255, b: n & 255 };
}

function drawDecorativeBlobs(shapes, accentColor, corner = "topRight") {
  // A bold, layered color-block accent — two overlapping, mostly-opaque
  // rectangles in two shades of the accent color. Built only from
  // rectangles and solid fills, both proven reliable today, rather than
  // gambling on unconfirmed shape types or rotation.
  const lightShade = shadeColor(accentColor, 0.35);
  const darkShade = shadeColor(accentColor, -0.25);

  const specs =
    corner === "topRight"
      ? [
          { left: SLIDE.width - 260, top: -40, width: 300, height: 300, color: lightShade, transparency: 0.15 },
          { left: SLIDE.width - 150, top: 40, width: 170, height: 170, color: darkShade, transparency: 0.05 },
        ]
      : [
          { left: -80, top: SLIDE.height - 260, width: 280, height: 280, color: lightShade, transparency: 0.2 },
          { left: 30, top: SLIDE.height - 150, width: 150, height: 150, color: accentColor, transparency: 0.1 },
        ];

  specs.forEach((s, idx) => {
    const block = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
    block.left = s.left;
    block.top = s.top;
    block.width = s.width;
    block.height = s.height;
    block.name = `GeminiSlides_Accent${corner}${idx}`;
    block.fill.setSolidColor(s.color);
    block.fill.transparency = s.transparency;
    block.lineFormat.visible = false;
  });
}

const TEMPLATES = {
  // --- Minimal: colored top bar, dark title, short accent rule, plain bullets.
  minimal(shapes, slideData, { isTitleSlide, slideNumber, accentColor, skipBackground }) {
    if (isTitleSlide && !skipBackground) {
      drawDecorativeBlobs(shapes, accentColor, "topRight");
      drawDecorativeBlobs(shapes, accentColor, "bottomLeft");
    }

    if (!(isTitleSlide && skipBackground)) {
      const accentBar = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
      accentBar.left = 0;
      accentBar.top = 0;
      accentBar.width = SLIDE.width;
      accentBar.height = isTitleSlide ? SLIDE.accentBarHeight * 2 : SLIDE.accentBarHeight;
      accentBar.name = "GeminiSlides_AccentBar";
      flatFill(accentBar, accentColor);
    }

    if (isTitleSlide) {
      drawCenteredTitle(shapes, slideData, accentColor, COLORS.dark);
    } else {
      const titleBox = shapes.addTextBox(slideData.title);
      Object.assign(titleBox, { left: SLIDE.left, top: SLIDE.titleTop, width: SLIDE.contentWidth, height: SLIDE.titleHeight });
      titleBox.name = "GeminiSlides_Title";
      titleBox.textFrame.textRange.font.size = 28;
      titleBox.textFrame.textRange.font.bold = true;
      titleBox.textFrame.textRange.font.color = COLORS.dark;

      const rule = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
      Object.assign(rule, { left: SLIDE.left, top: SLIDE.ruleTop, width: SLIDE.ruleWidth, height: SLIDE.ruleHeight });
      rule.name = "GeminiSlides_Rule";
      flatFill(rule, accentColor);

      drawPlainBullets(shapes, slideData.bullets, COLORS.body);
      drawBadge(shapes, slideNumber, {
        left: SLIDE.width - SLIDE.left - SLIDE.badgeSize + 20,
        top: SLIDE.height - SLIDE.badgeSize - 25,
        fillColor: accentColor,
        textColor: COLORS.white,
      });
    }
  },

  // --- Bold Panel: full-height colored sidebar holding the title.
  bold(shapes, slideData, { isTitleSlide, slideNumber, accentColor }) {
    const panelWidth = isTitleSlide ? SLIDE.width : 300;

    const panel = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
    Object.assign(panel, { left: 0, top: 0, width: panelWidth, height: SLIDE.height });
    panel.name = "GeminiSlides_Panel";
    flatFill(panel, accentColor);

    if (isTitleSlide) {
      const titleBox = shapes.addTextBox(slideData.title);
      Object.assign(titleBox, { left: 80, top: 210, width: SLIDE.width - 160, height: 100 });
      titleBox.name = "GeminiSlides_Title";
      titleBox.textFrame.textRange.font.size = 40;
      titleBox.textFrame.textRange.font.bold = true;
      titleBox.textFrame.textRange.font.color = COLORS.white;
      titleBox.textFrame.textRange.paragraphFormat.horizontalAlignment = PowerPoint.ParagraphHorizontalAlignment.center;

      const subtitle = slideData.bullets[0] || "";
      if (subtitle) {
        const subtitleBox = shapes.addTextBox(subtitle);
        Object.assign(subtitleBox, { left: 80, top: 310, width: SLIDE.width - 160, height: 50 });
        subtitleBox.name = "GeminiSlides_Subtitle";
        subtitleBox.textFrame.textRange.font.size = 20;
        subtitleBox.textFrame.textRange.font.italic = true;
        subtitleBox.textFrame.textRange.font.color = "#E8E6FF";
        subtitleBox.textFrame.textRange.paragraphFormat.horizontalAlignment = PowerPoint.ParagraphHorizontalAlignment.center;
      }
      return;
    }

    const titleBox = shapes.addTextBox(slideData.title);
    Object.assign(titleBox, { left: 24, top: 220, width: panelWidth - 48, height: 140 });
    titleBox.name = "GeminiSlides_Title";
    titleBox.textFrame.textRange.font.size = 24;
    titleBox.textFrame.textRange.font.bold = true;
    titleBox.textFrame.textRange.font.color = COLORS.white;

    const contentLeft = panelWidth + 40;
    const contentWidth = SLIDE.width - contentLeft - 40;
    drawPlainBullets(shapes, slideData.bullets, COLORS.body, { left: contentLeft, width: contentWidth, top: 60 });
    drawBadge(shapes, slideNumber, {
      left: SLIDE.width - SLIDE.badgeSize - 30,
      top: SLIDE.height - SLIDE.badgeSize - 25,
      fillColor: accentColor,
      textColor: COLORS.white,
    });
  },

  // --- Card Layout: each bullet rendered as its own tinted card.
  card(shapes, slideData, { isTitleSlide, slideNumber, accentColor, skipBackground }) {
    if (isTitleSlide) {
      if (!skipBackground) {
        drawDecorativeBlobs(shapes, accentColor, "topRight");
        drawDecorativeBlobs(shapes, accentColor, "bottomLeft");
      }
      const rule = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
      Object.assign(rule, { left: SLIDE.width / 2 - 60, top: 270, width: 120, height: 4 });
      rule.name = "GeminiSlides_Rule";
      flatFill(rule, accentColor);
      drawCenteredTitle(shapes, slideData, accentColor, COLORS.dark);
      return;
    }

    const titleBox = shapes.addTextBox(slideData.title);
    Object.assign(titleBox, { left: SLIDE.left, top: 45, width: SLIDE.contentWidth, height: 55 });
    titleBox.name = "GeminiSlides_Title";
    titleBox.textFrame.textRange.font.size = 26;
    titleBox.textFrame.textRange.font.bold = true;
    titleBox.textFrame.textRange.font.color = COLORS.dark;

    const bullets = slideData.bullets.length ? slideData.bullets : [""];
    const gap = 10;
    const top0 = 115;
    const bottom = 470;
    const cardHeight = Math.min(78, (bottom - top0 - gap * (bullets.length - 1)) / bullets.length);

    bullets.forEach((bullet, idx) => {
      const cardTop = top0 + idx * (cardHeight + gap);

      const card = shapes.addGeometricShape(PowerPoint.GeometricShapeType.roundedRectangle);
      Object.assign(card, { left: SLIDE.left, top: cardTop, width: SLIDE.contentWidth, height: cardHeight });
      card.name = `GeminiSlides_Card${idx}`;
      card.fill.setSolidColor(COLORS.cardBg);
      card.lineFormat.visible = false;

      const stripe = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
      Object.assign(stripe, { left: SLIDE.left, top: cardTop, width: 6, height: cardHeight });
      stripe.name = `GeminiSlides_CardStripe${idx}`;
      flatFill(stripe, accentColor);

      const textBox = shapes.addTextBox(bullet);
      Object.assign(textBox, {
        left: SLIDE.left + 24,
        top: cardTop,
        width: SLIDE.contentWidth - 48,
        height: cardHeight,
      });
      textBox.name = `GeminiSlides_CardText${idx}`;
      textBox.textFrame.textRange.font.size = 15;
      textBox.textFrame.textRange.font.color = COLORS.body;
      textBox.textFrame.wordWrap = true;
      textBox.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
    });

    drawBadge(shapes, slideNumber, {
      left: SLIDE.width - SLIDE.left - SLIDE.badgeSize + 20,
      top: SLIDE.height - SLIDE.badgeSize - 12,
      fillColor: accentColor,
      textColor: COLORS.white,
    });
  },

  // --- Dark Mode: dark full-slide background, light text, accent pops.
  dark(shapes, slideData, { isTitleSlide, slideNumber, accentColor, skipBackground }) {
    if (!(isTitleSlide && skipBackground)) {
      const bg = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
      Object.assign(bg, { left: 0, top: 0, width: SLIDE.width, height: SLIDE.height });
      bg.name = "GeminiSlides_Background";
      flatFill(bg, COLORS.darkBg);
    }

    if (isTitleSlide) {
      if (!skipBackground) {
        drawDecorativeBlobs(shapes, accentColor, "topRight");
        drawDecorativeBlobs(shapes, accentColor, "bottomLeft");
      }
      drawCenteredTitle(shapes, slideData, accentColor, COLORS.white);
      return;
    }

    const titleBox = shapes.addTextBox(slideData.title);
    Object.assign(titleBox, { left: SLIDE.left, top: SLIDE.titleTop, width: SLIDE.contentWidth, height: SLIDE.titleHeight });
    titleBox.name = "GeminiSlides_Title";
    titleBox.textFrame.textRange.font.size = 28;
    titleBox.textFrame.textRange.font.bold = true;
    titleBox.textFrame.textRange.font.color = COLORS.white;

    const rule = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
    Object.assign(rule, { left: SLIDE.left, top: SLIDE.ruleTop, width: SLIDE.ruleWidth, height: SLIDE.ruleHeight });
    rule.name = "GeminiSlides_Rule";
    flatFill(rule, accentColor);

    drawPlainBullets(shapes, slideData.bullets, COLORS.darkBody);
    drawBadge(shapes, slideNumber, {
      left: SLIDE.width - SLIDE.left - SLIDE.badgeSize + 20,
      top: SLIDE.height - SLIDE.badgeSize - 25,
      fillColor: accentColor,
      textColor: COLORS.darkBg,
    });
  },
};

function drawCenteredTitle(shapes, slideData, accentColor, titleColor) {
  const titleBox = shapes.addTextBox(slideData.title);
  Object.assign(titleBox, { left: 80, top: 190, width: SLIDE.width - 160, height: 110 });
  titleBox.name = "GeminiSlides_Title";
  titleBox.textFrame.textRange.font.size = 40;
  titleBox.textFrame.textRange.font.bold = true;
  titleBox.textFrame.textRange.font.color = titleColor;
  titleBox.textFrame.textRange.paragraphFormat.horizontalAlignment = PowerPoint.ParagraphHorizontalAlignment.center;

  const subtitle = slideData.bullets[0] || "";
  if (subtitle) {
    const subtitleBox = shapes.addTextBox(subtitle);
    Object.assign(subtitleBox, { left: 80, top: 300, width: SLIDE.width - 160, height: 50 });
    subtitleBox.name = "GeminiSlides_Subtitle";
    subtitleBox.textFrame.textRange.font.size = 20;
    subtitleBox.textFrame.textRange.font.italic = true;
    subtitleBox.textFrame.textRange.font.color = accentColor;
    subtitleBox.textFrame.textRange.paragraphFormat.horizontalAlignment = PowerPoint.ParagraphHorizontalAlignment.center;
  }
}

function drawPlainBullets(shapes, bullets, color, overrides = {}) {
  const bodyText = bullets.map((b) => `•  ${b}`).join("\n");
  if (!bodyText) return;
  const bodyBox = shapes.addTextBox(bodyText);
  bodyBox.left = overrides.left ?? SLIDE.left;
  bodyBox.top = overrides.top ?? SLIDE.bodyTop;
  bodyBox.width = overrides.width ?? SLIDE.contentWidth;
  bodyBox.height = overrides.height ?? SLIDE.bodyHeight;
  bodyBox.name = "GeminiSlides_Body";
  bodyBox.textFrame.textRange.font.size = 18;
  bodyBox.textFrame.textRange.font.color = color;
  bodyBox.textFrame.wordWrap = true;
}

async function buildDeckInPowerPoint(deck, { includeNotes, template, accentColor }, onSlideDone) {
  const addOptions = await findBlankLayoutOptions();

  for (let i = 0; i < deck.length; i++) {
    // eslint-disable-next-line no-await-in-loop
    await addOneSlide(deck[i], i, addOptions, { includeNotes, template, accentColor });
    onSlideDone(i);

    // PowerPoint on the web appears to need a moment to fully persist a
    // slide's content server-side before it's safe to start the next one —
    // a resolved await only confirms the client-side call returned, not
    // that the backend has caught up. This pause is a deliberate, brute
    // force mitigation for that lag.
    if (i < deck.length - 1) {
      // eslint-disable-next-line no-await-in-loop
      await new Promise((resolve) => setTimeout(resolve, 700));
    }
  }
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
  const template = els.template.value;
  const accentColor = els.accentColor.value;
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

    await buildDeckInPowerPoint(deck, { includeNotes, template, accentColor }, (i) => markProgress(i, "done"));
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
