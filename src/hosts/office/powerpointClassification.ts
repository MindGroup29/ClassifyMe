/*
 * PowerPoint-specific classification behavior.
 * The MVP marks only the slides that already exist when the user clicks a classification.
 */

/* global Office, PowerPoint, console */

import {
  CLASSIFICATION_LEVELS,
  CLASSIFICATION_PROPERTY_NAMES,
  CLASSIFICATION_TOOL_NAME,
  ClassificationCode,
  ClassificationLevel,
  getClassificationMarkingText,
  POWERPOINT_FOOTER_SHAPE_NAME,
} from "../../core/classificationConstants";

const FOOTER_MARGIN = 24;
const FOOTER_HEIGHT = 22;
const FOOTER_FONT_SIZE = 8;
const DEFAULT_SLIDE_WIDTH = 960;
const DEFAULT_SLIDE_HEIGHT = 540;
const METADATA_REQUIREMENT_SET_VERSION = "1.7";

// Internal technical state; the task pane does not automatically reload classification.
export type PowerPointClassificationState =
  | { status: "classified"; level: ClassificationCode }
  | { status: "unclassified" }
  | { status: "indeterminate" };

export async function applyPowerPointClassification(level: ClassificationLevel): Promise<number> {
  const slideCount = await applySlideMarkings(level);
  await tryUpdatePresentationMetadata(level);
  return slideCount;
}

async function applySlideMarkings(level: ClassificationLevel): Promise<number> {
  return PowerPoint.run(async (context) => {
    const presentation = context.presentation;
    const slides = presentation.slides;

    slides.load("items");
    await context.sync();

    slides.items.forEach((slide) => {
      slide.shapes.load("items/name");
    });
    await context.sync();

    slides.items.forEach((slide) => {
      removeExistingFooter(slide);
      addClassificationFooter(slide, level);
    });

    await context.sync();

    return slides.items.length;
  });
}

async function tryUpdatePresentationMetadata(level: ClassificationLevel): Promise<void> {
  try {
    if (!supportsPresentationMetadata()) {
      return;
    }

    // Commit footers first, in a separate run: a metadata error must not undo visual marking.
    await PowerPoint.run(async (context) => {
      const customProperties = context.presentation.properties.customProperties;
      const updatedAt = new Date().toISOString();

      // PowerPointApi 1.7 add() creates or updates by case-insensitive key, without duplicates.
      customProperties.add(CLASSIFICATION_PROPERTY_NAMES.level, level.code);
      customProperties.add(CLASSIFICATION_PROPERTY_NAMES.label, level.label);
      customProperties.add(CLASSIFICATION_PROPERTY_NAMES.updatedAt, updatedAt);
      customProperties.add(CLASSIFICATION_PROPERTY_NAMES.tool, CLASSIFICATION_TOOL_NAME);
      await context.sync();
    });
  } catch (error) {
    console.warn(
      "[ClassifyMe][PowerPoint] Impossible d'enregistrer les métadonnées ; les pieds de page sont conservés.",
      error
    );
  }
}

export async function readPowerPointClassification(): Promise<PowerPointClassificationState> {
  try {
    if (!supportsPresentationMetadata()) {
      return { status: "indeterminate" };
    }

    return await PowerPoint.run(async (context): Promise<PowerPointClassificationState> => {
      const customProperties = context.presentation.properties.customProperties;
      customProperties.load("items/key,items/value");
      await context.sync();

      // Include label and date when detecting partial metadata; shapes are never a source of truth.
      const propertyNames = Object.keys(CLASSIFICATION_PROPERTY_NAMES).map((name) =>
        CLASSIFICATION_PROPERTY_NAMES[name].toLowerCase()
      );
      const classificationProperties = customProperties.items.filter(
        (property) => propertyNames.indexOf(property.key.toLowerCase()) !== -1
      );
      if (classificationProperties.length === 0) {
        return { status: "unclassified" };
      }

      const levelProperty = classificationProperties.find(
        (property) =>
          property.key.toLowerCase() === CLASSIFICATION_PROPERTY_NAMES.level.toLowerCase()
      );
      const toolProperty = classificationProperties.find(
        (property) =>
          property.key.toLowerCase() === CLASSIFICATION_PROPERTY_NAMES.tool.toLowerCase()
      );
      const level = CLASSIFICATION_LEVELS.find(
        (candidate) => candidate.code === levelProperty?.value
      );
      if (toolProperty?.value === CLASSIFICATION_TOOL_NAME && level) {
        return { status: "classified", level: level.code };
      }

      return { status: "indeterminate" };
    });
  } catch (error) {
    console.warn("[ClassifyMe][PowerPoint] Impossible de relire les métadonnées.", error);
    return { status: "indeterminate" };
  }
}

function supportsPresentationMetadata(): boolean {
  const supported = Office.context.requirements.isSetSupported(
    "PowerPointApi",
    METADATA_REQUIREMENT_SET_VERSION
  );
  if (!supported) {
    console.warn(
      "[ClassifyMe][PowerPoint] PowerPointApi 1.7 indisponible : métadonnées inaccessibles, marquage visuel conservé."
    );
  }
  return supported;
}

function removeExistingFooter(slide: PowerPoint.Slide): void {
  slide.shapes.items
    .filter((shape) => shape.name === POWERPOINT_FOOTER_SHAPE_NAME)
    .forEach((shape) => shape.delete());
}

function addClassificationFooter(slide: PowerPoint.Slide, level: ClassificationLevel): void {
  const footer = slide.shapes.addTextBox(getClassificationMarkingText(level), {
    left: FOOTER_MARGIN,
    top: DEFAULT_SLIDE_HEIGHT - FOOTER_MARGIN - FOOTER_HEIGHT,
    width: DEFAULT_SLIDE_WIDTH - FOOTER_MARGIN * 2,
    height: FOOTER_HEIGHT,
  });

  footer.name = POWERPOINT_FOOTER_SHAPE_NAME;
  footer.fill.setSolidColor(level.bannerBackground);
  footer.lineFormat.visible = false;

  // Keep formatting limited to stable PowerPointApi 1.4 properties. More advanced
  // slide-size and autosize APIs can raise GeneralException on some desktop builds.
  footer.textFrame.leftMargin = 6;
  footer.textFrame.rightMargin = 6;
  footer.textFrame.topMargin = 2;
  footer.textFrame.bottomMargin = 2;
  footer.textFrame.wordWrap = true;
  footer.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;

  footer.textFrame.textRange.font.name = "Segoe UI";
  footer.textFrame.textRange.font.size = FOOTER_FONT_SIZE;
  footer.textFrame.textRange.font.color = level.bannerColor;
  footer.textFrame.textRange.font.bold = level.code === "CONFIDENTIEL" || level.code === "SECRET";
  footer.textFrame.textRange.paragraphFormat.horizontalAlignment =
    PowerPoint.ParagraphHorizontalAlignment.center;
}
