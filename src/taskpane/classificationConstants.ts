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
export const POWERPOINT_FOOTER_SHAPE_NAME = "ClassifyMeFooter";

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
    bannerTitle: "",
    bannerText: "",
    bannerColor: "#1f2933",
    bannerBackground: "#baf1cf",
  },
  {
    code: "RESTREINT",
    label: "Restreint",
    description: "Information reserved for company employees.",
    bannerTitle: "Classification: RESTREINT",
    bannerText: "Ce document est destine a un usage restreint a l'entreprise.",
    bannerColor: "#1f2933",
    bannerBackground: "#e8f1fb",
  },
  {
    code: "CONFIDENTIEL",
    label: "Confidentiel",
    description: "Information whose uncontrolled disclosure may harm the company.",
    bannerTitle: "Classification: CONFIDENTIEL",
    bannerText:
      "Ce document contient des informations confidentielles. Sa diffusion doit etre limitee aux personnes autorisees.",
    bannerColor: "#ffffff",
    bannerBackground: "#9f1239",
  },
  {
    code: "SECRET",
    label: "Secret",
    description: "Critical information requiring strict control before sharing.",
    bannerTitle: "Classification: SECRET",
    bannerText:
      "Ce document contient des informations secretes. Sa diffusion, sa copie et son transfert doivent etre strictement maitrises.",
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
