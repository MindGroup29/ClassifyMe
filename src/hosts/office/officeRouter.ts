/*
 * Office host router for ClassifyMe Office.
 * The UI calls this module once; the router keeps Word, Excel and PowerPoint branching out of the task pane.
 */

/* global Office */

import { ClassificationLevel } from "../../core/classificationConstants";
import { applyExcelClassification } from "./excelClassification";
import { applyPowerPointClassification } from "./powerpointClassification";
import { applyWordClassification } from "./wordClassification";

export function isSupportedOfficeHost(host: Office.HostType | undefined): boolean {
  return (
    host === Office.HostType.Word ||
    host === Office.HostType.PowerPoint ||
    host === Office.HostType.Excel
  );
}

export function getOfficeHostDocumentName(host: Office.HostType | undefined): string {
  if (host === Office.HostType.PowerPoint) {
    return "presentation";
  }

  if (host === Office.HostType.Excel) {
    return "workbook";
  }

  return "document";
}

export async function applyOfficeClassification(
  host: Office.HostType | undefined,
  level: ClassificationLevel
): Promise<string> {
  if (host === Office.HostType.Word) {
    await applyWordClassification(level);
    return `Classification ${level.label} applied.`;
  }

  if (host === Office.HostType.PowerPoint) {
    const slideCount = await applyPowerPointClassification(level);
    return `Classification ${level.label} applied to ${slideCount} slide(s).`;
  }

  if (host === Office.HostType.Excel) {
    const result = await applyExcelClassification(level);
    const metadataMessage = result.metadataSaved
      ? ""
      : " Workbook metadata could not be saved in this Excel environment.";

    return `Classification ${level.label} applied to ${result.worksheetCount} worksheet(s).${metadataMessage}`;
  }

  return "ClassifyMe Office supports Word, PowerPoint and Excel only in this MVP.";
}
