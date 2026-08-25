export type ClassificationCode = "PUBLIC" | "RESTREINT" | "CONFIDENTIEL" | "SECRET";

export interface ClassificationLevel {
  code: ClassificationCode;
  label: string;
  description: string;
  bannerTitle: string;
  bannerText: string;
  bannerColor: string;
  bannerBackground: string;
}

export const CLASSIFICATION_TOOL_NAME = "ClassifyMe";
export const CLASSIFICATION_BANNER_TAG = "ClassifyMe.ClassificationBanner";
export const OUTLOOK_BANNER_START_MARKER = "<!-- ClassifyMe:BannerStart -->";
export const OUTLOOK_BANNER_END_MARKER = "<!-- ClassifyMe:BannerEnd -->";
export const POWERPOINT_FOOTER_SHAPE_NAME = "ClassifyMeFooter";
export const EXCEL_BANNER_SHAPE_NAME = "ClassifyMeBanner";

export const CLASSIFICATION_PROPERTY_NAMES = {
  level: "ClassificationLevel",
  label: "ClassificationLabel",
  updatedAt: "ClassificationUpdatedAt",
  tool: "ClassificationTool",
};

export const CLASSIFICATION_LEVELS: ClassificationLevel[] = [
  {
    code: "PUBLIC",
    label: "Public",
    description: "Information intended for public sharing.",
    bannerTitle: "Classification : PUBLIC",
    bannerText: "Ce document peut être diffusé publiquement.",
    bannerColor: "#1f2933",
    bannerBackground: "#baf1cf",
  },
  {
    code: "RESTREINT",
    label: "Restreint",
    description: "Information reserved for company employees.",
    bannerTitle: "Classification : RESTREINT",
    bannerText: "Ce document est destiné à un usage restreint à l'entreprise.",
    bannerColor: "#1f2933",
    bannerBackground: "#e8f1fb",
  },
  {
    code: "CONFIDENTIEL",
    label: "Confidentiel",
    description: "Information whose uncontrolled disclosure may harm the company.",
    bannerTitle: "Classification : CONFIDENTIEL",
    bannerText:
      "Ce document contient des informations confidentielles. Sa diffusion doit être limitée aux personnes autorisées.",
    bannerColor: "#ffffff",
    bannerBackground: "#9f1239",
  },
  {
    code: "SECRET",
    label: "Secret",
    description: "Critical information requiring strict control before sharing.",
    bannerTitle: "Classification : SECRET",
    bannerText:
      "Ce document contient des informations secrètes. Sa diffusion, sa copie et son transfert doivent être strictement maitrisés.",
    bannerColor: "#ffffff",
    bannerBackground: "#581c87",
  },
];

export function getClassificationMarkingText(level: ClassificationLevel): string {
  if (!level.bannerTitle || !level.bannerText) {
    return `Classification: ${level.code}`;
  }

  return `${level.bannerTitle} - ${level.bannerText}`;
}
