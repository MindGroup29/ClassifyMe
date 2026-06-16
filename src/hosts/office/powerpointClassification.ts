/*
 * PowerPoint-specific classification behavior.
 * The MVP marks only the slides that already exist when the user clicks a classification.
 */

/* global PowerPoint */

import {
  ClassificationLevel,
  getClassificationMarkingText,
  POWERPOINT_FOOTER_SHAPE_NAME,
} from "../../core/classificationConstants";

const FOOTER_MARGIN = 24;
const FOOTER_HEIGHT = 22;
const FOOTER_FONT_SIZE = 8;
const DEFAULT_SLIDE_WIDTH = 960;
const DEFAULT_SLIDE_HEIGHT = 540;

export async function applyPowerPointClassification(level: ClassificationLevel): Promise<number> {
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
