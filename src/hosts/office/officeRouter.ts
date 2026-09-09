/*
 * Office host router for ClassifyMe.
 * The UI calls this module once; the router keeps host-specific branching out of the task pane.
 */

/* global Office */

import { ClassificationLevel } from "../../core/classificationConstants";
import {
  applyOutlookClassification,
  OutlookClassificationOptions,
} from "../outlook/outlookClassification";
import { applyExcelClassification } from "./excelClassification";
import { applyPowerPointClassification } from "./powerpointClassification";
import { applyWordClassification } from "./wordClassification";

export function isSupportedOfficeHost(host: Office.HostType | undefined): boolean {
  return (
    host === Office.HostType.Word ||
    host === Office.HostType.PowerPoint ||
    host === Office.HostType.Excel ||
    host === Office.HostType.Outlook
  );
}

export function getOfficeHostDocumentName(host: Office.HostType | undefined): string {
  if (host === Office.HostType.Outlook) {
    return "message Outlook";
  }

  if (host === Office.HostType.PowerPoint) {
    return "presentation";
  }

  if (host === Office.HostType.Excel) {
    return "classeur";
  }

  return "document";
}

export async function applyOfficeClassification(
  host: Office.HostType | undefined,
  level: ClassificationLevel,
  options: OutlookClassificationOptions = { addSubjectPrefix: false }
): Promise<string> {
  if (host === Office.HostType.Word) {
    await applyWordClassification(level);
    return `Classification ${level.label} appliquée.`;
  }

  if (host === Office.HostType.PowerPoint) {
    const slideCount = await applyPowerPointClassification(level);
    return `Classification ${level.label} appliquée a ${slideCount} slide(s).`;
  }

  if (host === Office.HostType.Excel) {
    const result = await applyExcelClassification(level);
    const metadataMessage = result.metadataSaved
      ? ""
      : " Les metadonnees du classeur n'ont pas pu etre enregistrees dans cet environnement Excel.";

    return `Classification ${level.label} appliquée a ${result.worksheetCount} feuille(s).${metadataMessage}`;
  }

  if (host === Office.HostType.Outlook) {
    const result = await applyOutlookClassification(level, options);
    const metadataMessage = result.metadataSaved
      ? ""
      : " Les métadonnées Outlook n'ont pas pu être enregistrées dans cet environnement.";
    const subjectMessage = result.subjectUpdated ? " Objet mis a jour." : "";

    const itemLabel = result.itemType === "appointment" ? "cette réunion" : "cet email";

    return `Classification ${level.label} appliquée à ${itemLabel}.${subjectMessage}${metadataMessage}`;
  }

  return "ClassifyMe prend en charge Word, PowerPoint, Excel et Outlook en mode composition uniquement dans ce MVP.";
}
