/*
 * ClassifyMe MVP for Word.
 * The add-in only applies a visible user-selected classification banner and simple document metadata.
 */

/* global document, Office, Word */

import {
  CLASSIFICATION_BANNER_TAG,
  CLASSIFICATION_LEVELS,
  CLASSIFICATION_PROPERTY_NAMES,
  CLASSIFICATION_TOOL_NAME,
  ClassificationCode,
  ClassificationLevel,
} from "./classificationConstants";

let selectedLevel: ClassificationLevel | undefined;

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    renderClassificationCards();
    updateSelectedLevel(undefined);

    document.getElementById("apply-classification").onclick = applySelectedClassification;
  }
});

function renderClassificationCards(): void {
  const levelsContainer = document.getElementById("classification-levels");
  levelsContainer.innerHTML = "";

  CLASSIFICATION_LEVELS.forEach((level) => {
    const card = document.createElement("button");
    card.type = "button";
    card.className = "classification-card";
    card.dataset.level = level.code;
    card.innerHTML = `
      <span class="classification-card__label">${level.label}</span>
      <span class="classification-card__code">${level.code}</span>
      <span class="classification-card__description">${level.description}</span>
    `;
    card.onclick = () => updateSelectedLevel(level);

    levelsContainer.appendChild(card);
  });
}

function updateSelectedLevel(level: ClassificationLevel | undefined): void {
  selectedLevel = level;

  const selectedLabel = document.getElementById("selected-level");
  const applyButton = document.getElementById("apply-classification") as any;

  selectedLabel.textContent = level
    ? `${level.label} (${level.code})`
    : "No classification selected";
  applyButton.disabled = !level;

  document.querySelectorAll(".classification-card").forEach((card) => {
    card.classList.toggle(
      "classification-card--selected",
      card.getAttribute("data-level") === level?.code
    );
  });
}

async function applySelectedClassification(): Promise<void> {
  if (!selectedLevel) {
    showStatus("Choose a classification level first.", "error");
    return;
  }

  setBusyState(true);
  showStatus("Applying classification...", "info");

  try {
    await Word.run(async (context) => {
      await applyDocumentBanner(context, selectedLevel);
      await updateDocumentMetadata(context, selectedLevel);
      await context.sync();
    });

    const publicMessage =
      selectedLevel.code === "PUBLIC"
        ? " Public documents do not receive a banner by default."
        : "";
    showStatus(`Classification ${selectedLevel.label} applied.${publicMessage}`, "success");
  } catch (error) {
    showStatus(`Unable to apply classification: ${getErrorMessage(error)}`, "error");
  } finally {
    setBusyState(false);
  }
}

async function applyDocumentBanner(
  context: Word.RequestContext,
  level: ClassificationLevel
): Promise<void> {
  const body = context.document.body;
  const existingBanners = body.contentControls.getByTag(CLASSIFICATION_BANNER_TAG);

  existingBanners.load("items");
  await context.sync();

  // PUBLIC has no visible banner in classification-rules.md, so an existing ClassifyMe banner is removed.
  if (level.code === "PUBLIC") {
    existingBanners.items.forEach((banner) => banner.delete(false));
    return;
  }

  existingBanners.items.forEach((banner) => banner.delete(false));
  createClassificationBannerTable(body, level);
}

function createClassificationBannerTable(
  body: Word.Body,
  level: ClassificationLevel
): Word.ContentControl {
  const bannerText = `${level.bannerTitle} - ${level.bannerText}`;
  const table = body.insertTable(1, 1, Word.InsertLocation.start, [[bannerText]]);
  const banner = table.insertContentControl();
  const cell = table.getCell(0, 0);
  const paragraph = cell.body.paragraphs.getFirst();

  banner.tag = CLASSIFICATION_BANNER_TAG;
  banner.title = "ClassifyMe classification banner";
  banner.appearance = Word.ContentControlAppearance.hidden;
  banner.cannotEdit = false;
  banner.cannotDelete = false;

  // A one-cell table gives Word a real cell background. This is more reliable than
  // paragraph shading for a persistent banner background across Word versions.
  table.alignment = Word.Alignment.left;
  table.horizontalAlignment = Word.Alignment.left;
  table.styleBuiltIn = Word.BuiltInStyleName.plainTable1;
  table.autoFitWindow();

  table.font.bold = true;
  table.font.color = level.bannerColor;
  cell.shadingColor = level.bannerBackground;
  cell.horizontalAlignment = Word.Alignment.left;
  paragraph.font.bold = true;
  paragraph.font.color = level.bannerColor;
  paragraph.spaceAfter = 6;
  paragraph.spaceBefore = 6;

  ["Top", "Left", "Bottom", "Right", "InsideHorizontal", "InsideVertical"].forEach(
    (borderLocation) => {
      table.getBorder(borderLocation as any).type = Word.BorderType.none;
    }
  );

  return banner;
}

async function updateDocumentMetadata(
  context: Word.RequestContext,
  level: ClassificationLevel
): Promise<void> {
  const customProperties = context.document.properties.customProperties;
  const updatedAt = new Date().toISOString();

  // Word custom properties are updated through the queued Office.js context, then committed by context.sync().
  await setCustomProperty(
    context,
    customProperties,
    CLASSIFICATION_PROPERTY_NAMES.level,
    level.code
  );
  await setCustomProperty(
    context,
    customProperties,
    CLASSIFICATION_PROPERTY_NAMES.label,
    level.label
  );
  await setCustomProperty(
    context,
    customProperties,
    CLASSIFICATION_PROPERTY_NAMES.updatedAt,
    updatedAt
  );
  await setCustomProperty(
    context,
    customProperties,
    CLASSIFICATION_PROPERTY_NAMES.tool,
    CLASSIFICATION_TOOL_NAME
  );
}

async function setCustomProperty(
  context: Word.RequestContext,
  customProperties: Word.CustomPropertyCollection,
  propertyName: string,
  propertyValue: ClassificationCode | string
): Promise<void> {
  const existingProperty = customProperties.getItemOrNullObject(propertyName);

  existingProperty.load("value");
  await context.sync();

  if (existingProperty.isNullObject) {
    customProperties.add(propertyName, propertyValue);
    return;
  }

  existingProperty.value = propertyValue;
}

function setBusyState(isBusy: boolean): void {
  const applyButton = document.getElementById("apply-classification") as any;

  applyButton.disabled = isBusy || !selectedLevel;
  applyButton.textContent = isBusy ? "Applying..." : "Apply classification";
}

function showStatus(message: string, type: "info" | "success" | "error"): void {
  const status = document.getElementById("status-message");

  status.textContent = message;
  status.className = `status-message status-message--${type}`;
}

function getErrorMessage(error: unknown): string {
  if (error instanceof Error) {
    return error.message;
  }

  return String(error);
}
