export type ClassificationCode = "PUBLIC" | "RESTREINT" | "CONFIDENTIEL" | "SECRET";

export interface ClassificationLevel {
  code: ClassificationCode;
  label: string;
  description: string;
  bannerTitle: string;
  bannerText: string;
  // Outlook meetings use a context-specific message while sharing the same level and colors.
  meetingBannerText: string;
  bannerColor: string;
  bannerBackground: string;
}

export const CLASSIFICATION_TOOL_NAME = "ClassifyMe";
export const CLASSIFICATION_BANNER_TAG = "ClassifyMe.ClassificationBanner";
// A real HTML element identifier is more resilient than HTML comments when Outlook rewrites a draft body.
export const OUTLOOK_BANNER_ELEMENT_ID = "classifyme-classification-banner";
// Outlook on the web can add this prefix to id and class attributes while rendering email HTML.
export const OUTLOOK_WEB_HTML_ID_PREFIX = "x_";
// These markers identify the first pilot format and are kept only to clean up existing drafts.
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
    description: "Information destinée à être partagée publiquement.",
    bannerTitle: "Classification : PUBLIC",
    bannerText: "Ce document peut être diffusé publiquement.",
    meetingBannerText:
      "Cette réunion contient des informations destinées à être partagées publiquement.",
    bannerColor: "#1f2933",
    bannerBackground: "#baf1cf",
  },
  {
    code: "RESTREINT",
    label: "Restreint",
    description: "Information réservée aux collaborateurs de l’entreprise.",
    bannerTitle: "Classification : RESTREINT",
    bannerText: "Ce document est destiné à un usage restreint à l'entreprise.",
    meetingBannerText:
      "Cette réunion contient des informations réservées aux collaborateurs de l'entreprise.",
    bannerColor: "#1f2933",
    bannerBackground: "#e8f1fb",
  },
  {
    code: "CONFIDENTIEL",
    label: "Confidentiel",
    description:
      "Information dont la divulgation non maîtrisée peut porter préjudice à l’entreprise.",
    bannerTitle: "Classification : CONFIDENTIEL",
    bannerText:
      "Ce document contient des informations confidentielles. Sa diffusion doit être limitée aux personnes autorisées.",
    meetingBannerText:
      "Cette réunion contient des informations confidentielles. La diffusion de son contenu doit être limitée aux destinataires autorisés.",
    bannerColor: "#ffffff",
    bannerBackground: "#9f1239",
  },
  {
    code: "SECRET",
    label: "Secret",
    description: "Information critique nécessitant un contrôle strict avant tout partage.",
    bannerTitle: "Classification : SECRET",
    bannerText:
      "Ce document contient des informations secrètes. Sa diffusion, sa copie et son transfert doivent être strictement maitrisés.",
    meetingBannerText:
      "Cette réunion contient des informations secrètes. Sa diffusion, sa copie et son transfert doivent être strictement maîtrisés.",
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
