/*
 * ClassifyMe MVP task pane.
 * Host-specific Office.js behavior is delegated to the ClassifyMe Office router.
 */

/* global document, Office, HTMLInputElement */

import { CLASSIFICATION_LEVELS, ClassificationLevel } from "../../core/classificationConstants";
import {
  applyOfficeClassification,
  getOfficeHostDocumentName,
  isSupportedOfficeHost,
} from "../../hosts/office/officeRouter";

let activeHost: Office.HostType | undefined;

Office.onReady((info) => {
  activeHost = info.host;
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";

  if (isSupportedOfficeHost(activeHost)) {
    updateHostSpecificUi(activeHost);
    renderClassificationCards();
    updateSelectedLevel(undefined);
    showStatus(
      `Choisissez un niveau pour l'appliquer sur ce ${getOfficeHostDocumentName(activeHost)}.`,
      "info"
    );
    return;
  }

  updateSelectedLevel(undefined);
  showStatus(
    "ClassifyMe prend en charge Word, PowerPoint, Excel et Outlook en mode composition uniquement dans ce MVP.",
    "error"
  );
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
  showStatus(`Application de la classification ${level.label}...`, "info");

  try {
    const message = await applyOfficeClassification(activeHost, level, {
      addSubjectPrefix: shouldAddSubjectPrefix(),
    });
    showStatus(message, isSupportedOfficeHost(activeHost) ? "success" : "error");
  } catch (error) {
    showStatus(`Impossible d'appliquer la classification : ${getErrorMessage(error)}`, "error");
  } finally {
    setBusyState(false);
  }
}

function updateHostSpecificUi(host: Office.HostType | undefined): void {
  const subtitle = document.getElementById("host-subtitle");
  const outlookOption = document.getElementById("outlook-subject-prefix-option");

  if (host === Office.HostType.Outlook) {
    subtitle.textContent = "Classifiez cet email ou cette réunion avant l'envoi.";
    outlookOption.style.display = "flex";
    return;
  }

  subtitle.textContent = "Classifiez ce document avant le partage.";
  outlookOption.style.display = "none";
}

function shouldAddSubjectPrefix(): boolean {
  if (activeHost !== Office.HostType.Outlook) {
    return false;
  }

  const checkbox = document.getElementById("add-subject-prefix") as HTMLInputElement | null;

  return Boolean(checkbox?.checked);
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
