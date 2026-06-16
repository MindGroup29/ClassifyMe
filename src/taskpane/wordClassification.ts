/*
 * Word-specific classification behavior.
 * This preserves the existing Word MVP logic while keeping host-specific code out of the task pane UI.
 */

/* global Word */

import {
  CLASSIFICATION_BANNER_TAG,
  CLASSIFICATION_PROPERTY_NAMES,
  CLASSIFICATION_TOOL_NAME,
  ClassificationCode,
  ClassificationLevel,
  getClassificationMarkingText,
} from "./classificationConstants";

export async function applyWordClassification(level: ClassificationLevel): Promise<void> {
  await Word.run(async (context) => {
    await applyDocumentBanner(context, level);
    await updateDocumentMetadata(context, level);
    await context.sync();
  });
}

async function applyDocumentBanner(
  context: Word.RequestContext,
  level: ClassificationLevel
): Promise<void> {
  const body = context.document.body;
  const sections = context.document.sections;
  const existingBodyBanners = body.contentControls.getByTag(CLASSIFICATION_BANNER_TAG);

  existingBodyBanners.load("items");
  sections.load("items");
  await context.sync();

  const headerBannerCollections = sections.items.map((section) => {
    const header = section.getHeader(Word.HeaderFooterType.primary);
    const existingHeaderBanners = header.contentControls.getByTag(CLASSIFICATION_BANNER_TAG);

    existingHeaderBanners.load("items");
    return { header, existingHeaderBanners };
  });

  await context.sync();

  // Remove older ClassifyMe banners from the document body and from Word headers before recreating them.
  existingBodyBanners.items.forEach((banner) => banner.delete(false));
  headerBannerCollections.forEach(({ existingHeaderBanners }) => {
    existingHeaderBanners.items.forEach((banner) => banner.delete(false));
  });

  // PUBLIC has no visible banner in classification-rules.md, so an existing ClassifyMe banner is removed.
  // if (level.code === "PUBLIC") {
  //   return;
  // }

  // Word stores headers per section. Applying the banner to each primary header keeps
  // the classification visible throughout documents that contain section breaks.
  headerBannerCollections.forEach(({ header }) => {
    createClassificationBannerTable(header, level);
  });
}

function createClassificationBannerTable(
  body: Word.Body,
  level: ClassificationLevel
): Word.ContentControl {
  const table = body.insertTable(1, 1, Word.InsertLocation.start, [
    [getClassificationMarkingText(level)],
  ]);
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
