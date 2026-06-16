/*
 * ClassifyMe MVP task pane.
 * Host-specific Office.js behavior is delegated to Word and PowerPoint modules.
 */

/* global document, Office */

import { CLASSIFICATION_LEVELS, ClassificationLevel } from "./classificationConstants";
import { applyPowerPointClassification } from "./powerPointClassification";
import { applyWordClassification } from "./wordClassification";

let activeHost: Office.HostType | undefined;

Office.onReady((info) => {
  activeHost = info.host;
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";

  if (isSupportedHost(activeHost)) {
    renderClassificationCards();
    updateSelectedLevel(undefined);
    showStatus(`Choose a level to apply it to the ${getHostDocumentName()}.`, "info");
    return;
  }

  updateSelectedLevel(undefined);
  showStatus("ClassifyMe supports Word and PowerPoint only in this MVP.", "error");
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
    card.onclick = () => applyClassification(level);

    levelsContainer.appendChild(card);
  });
}

function updateSelectedLevel(level: ClassificationLevel | undefined): void {
  const selectedLabel = document.getElementById("selected-level");
  const selectedSummary = document.querySelector(".selection-summary") as any;
  const selectedSummaryLabel = document.querySelector(".selection-summary__label") as any;

  selectedLabel.textContent = level
    ? `${level.label} (${level.code})`
    : "No classification selected";

  if (level) {
    selectedSummary.style.backgroundColor = level.bannerBackground;
    selectedSummary.style.borderColor = level.bannerBackground;
    selectedSummary.style.color = level.bannerColor;
    selectedLabel.style.color = level.bannerColor;
    selectedSummaryLabel.style.color = level.bannerColor;
  } else {
    selectedSummary.removeAttribute("style");
    selectedLabel.removeAttribute("style");
    selectedSummaryLabel.removeAttribute("style");
  }

  document.querySelectorAll(".classification-card").forEach((card) => {
    card.classList.toggle(
      "classification-card--selected",
      card.getAttribute("data-level") === level?.code
    );
  });
}

async function applyClassification(level: ClassificationLevel): Promise<void> {
  updateSelectedLevel(level);

  setBusyState(true);
  showStatus(`Applying ${level.label} classification...`, "info");

  try {
    if (activeHost === Office.HostType.Word) {
      await applyWordClassification(level);

      // const publicMessage =
      //   level.code === "PUBLIC" ? " Public documents do not receive a banner by default." : "";
      showStatus(`Classification ${level.label} applied.`, "success");
      return;
    }

    if (activeHost === Office.HostType.PowerPoint) {
      const slideCount = await applyPowerPointClassification(level);
      showStatus(`Classification ${level.label} applied to ${slideCount} slide(s).`, "success");
      return;
    }

    showStatus("ClassifyMe supports Word and PowerPoint only in this MVP.", "error");
  } catch (error) {
    showStatus(`Unable to apply classification: ${getErrorMessage(error)}`, "error");
  } finally {
    setBusyState(false);
  }
}

function setBusyState(isBusy: boolean): void {
  document.querySelectorAll(".classification-card").forEach((card) => {
    (card as any).disabled = isBusy;
    card.setAttribute("aria-disabled", String(isBusy));
  });
}

function showStatus(message: string, type: "info" | "success" | "error"): void {
  const status = document.getElementById("status-message");

  status.textContent = message;
  status.className = `status-message status-message--${type}`;
}

function getErrorMessage(error: unknown): string {
  if (error instanceof Error) {
    const officeError = error as Error & {
      code?: string;
      debugInfo?: {
        errorLocation?: string;
        message?: string;
      };
    };
    const details = [
      officeError.code,
      officeError.debugInfo?.errorLocation,
      officeError.debugInfo?.message,
    ].filter(Boolean);

    if (details.length > 0) {
      return `${error.message} (${details.join(" - ")})`;
    }

    return error.message;
  }

  return String(error);
}

function isSupportedHost(host: Office.HostType | undefined): boolean {
  return host === Office.HostType.Word || host === Office.HostType.PowerPoint;
}

function getHostDocumentName(): string {
  return activeHost === Office.HostType.PowerPoint ? "presentation" : "document";
}
