/*
 * Excel-specific classification behavior.
 * The MVP marks only the worksheets that already exist when the user clicks a classification.
 */

/* global Excel */

import {
  CLASSIFICATION_PROPERTY_NAMES,
  CLASSIFICATION_TOOL_NAME,
  ClassificationLevel,
  EXCEL_BANNER_SHAPE_NAME,
  getClassificationMarkingText,
} from "./classificationConstants";

const BANNER_LEFT = 12;
const BANNER_TOP = 12;
const BANNER_WIDTH = 450;
const BANNER_HEIGHT = 30;
const BANNER_FONT_SIZE = 10;

export interface ExcelClassificationResult {
  worksheetCount: number;
  metadataSaved: boolean;
}

export async function applyExcelClassification(
  level: ClassificationLevel
): Promise<ExcelClassificationResult> {
  const worksheetCount = await applyWorksheetMarkings(level);
  const metadataSaved = await tryUpdateWorkbookMetadata(level);

  return { worksheetCount, metadataSaved };
}

async function applyWorksheetMarkings(level: ClassificationLevel): Promise<number> {
  return Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;

    worksheets.load("items/name");
    await context.sync();

    if (worksheets.items.length === 0) {
      throw new Error("No worksheet found in this workbook.");
    }

    worksheets.items.forEach((worksheet) => {
      worksheet.shapes.load("items/name");
    });
    await context.sync();

    worksheets.items.forEach((worksheet) => {
      removeExistingBanner(worksheet);
      addClassificationBanner(worksheet, level);
      updateWorksheetFooter(worksheet, level);
    });

    await context.sync();

    return worksheets.items.length;
  });
}

function removeExistingBanner(worksheet: Excel.Worksheet): void {
  worksheet.shapes.items
    .filter((shape) => shape.name === EXCEL_BANNER_SHAPE_NAME)
    .forEach((shape) => shape.delete());
}

function addClassificationBanner(worksheet: Excel.Worksheet, level: ClassificationLevel): void {
  const banner = worksheet.shapes.addTextBox(getClassificationMarkingText(level));

  banner.name = EXCEL_BANNER_SHAPE_NAME;
  banner.left = BANNER_LEFT;
  banner.top = BANNER_TOP;
  banner.width = BANNER_WIDTH;
  banner.height = BANNER_HEIGHT;
  banner.placement = Excel.Placement.absolute;
  banner.fill.setSolidColor(level.bannerBackground);
  banner.lineFormat.visible = false;

  // Excel text boxes need explicit sizing and alignment so the banner remains
  // readable without changing worksheet cells or business content.
  banner.textFrame.leftMargin = 8;
  banner.textFrame.rightMargin = 8;
  banner.textFrame.topMargin = 3;
  banner.textFrame.bottomMargin = 3;
  banner.textFrame.horizontalAlignment = Excel.ShapeTextHorizontalAlignment.center;
  banner.textFrame.verticalAlignment = Excel.ShapeTextVerticalAlignment.middle;
  banner.textFrame.textRange.font.name = "Segoe UI";
  banner.textFrame.textRange.font.size = BANNER_FONT_SIZE;
  banner.textFrame.textRange.font.color = level.bannerColor;
  banner.textFrame.textRange.font.bold = true;
}

function updateWorksheetFooter(worksheet: Excel.Worksheet, level: ClassificationLevel): void {
  const headersFooters = worksheet.pageLayout.headersFooters;

  headersFooters.state = Excel.HeaderFooterState.default;
  headersFooters.defaultForAllPages.centerFooter = getClassificationMarkingText(level);
}

async function tryUpdateWorkbookMetadata(level: ClassificationLevel): Promise<boolean> {
  try {
    await Excel.run(async (context) => {
      const customProperties = context.workbook.properties.custom;
      const updatedAt = new Date().toISOString();

      // Excel CustomPropertyCollection.add creates or updates the property by key.
      customProperties.add(CLASSIFICATION_PROPERTY_NAMES.level, level.code);
      customProperties.add(CLASSIFICATION_PROPERTY_NAMES.label, level.label);
      customProperties.add(CLASSIFICATION_PROPERTY_NAMES.updatedAt, updatedAt);
      customProperties.add(CLASSIFICATION_PROPERTY_NAMES.tool, CLASSIFICATION_TOOL_NAME);

      await context.sync();
    });

    return true;
  } catch {
    return false;
  }
}
