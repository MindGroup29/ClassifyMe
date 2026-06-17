/*
 * Outlook-specific classification behavior.
 * The MVP only supports message compose mode: it updates the draft body and, optionally, the subject.
 */

/* global Office */

import {
  CLASSIFICATION_PROPERTY_NAMES,
  CLASSIFICATION_TOOL_NAME,
  ClassificationLevel,
  OUTLOOK_BANNER_END_MARKER,
  OUTLOOK_BANNER_START_MARKER,
} from "../../core/classificationConstants";

const CLASSIFYME_SUBJECT_PREFIX_PATTERN =
  /^\s*\[(PUBLIC|RESTREINT|RESTRAINT|CONFIDENTIEL|SECRET)\]\s*/i;

export interface OutlookClassificationOptions {
  addSubjectPrefix: boolean;
}

export interface OutlookClassificationResult {
  metadataSaved: boolean;
  subjectUpdated: boolean;
}

export async function applyOutlookClassification(
  level: ClassificationLevel,
  options: OutlookClassificationOptions
): Promise<OutlookClassificationResult> {
  const item = getComposeMessage();
  const currentHtmlBody = await getBodyHtml(item);
  const updatedHtmlBody = addOrReplaceBanner(currentHtmlBody, buildOutlookBannerHtml(level));

  await setBodyHtml(item, updatedHtmlBody);

  const subjectUpdated = options.addSubjectPrefix
    ? await applySubjectPrefix(item, level)
    : await removeSubjectPrefix(item);

  const metadataSaved = await trySaveCustomProperties(item, level);

  return { metadataSaved, subjectUpdated };
}

function getComposeMessage(): Office.MessageCompose {
  const mailbox = Office.context.mailbox;
  const item = mailbox?.item as Partial<Office.MessageCompose> | undefined;

  if (!item?.body || !item?.subject || typeof item.loadCustomPropertiesAsync !== "function") {
    throw new Error(
      "Cet element Outlook ne peut pas etre classifie. Ouvrez un message en mode composition."
    );
  }

  return item as Office.MessageCompose;
}

function buildOutlookBannerHtml(level: ClassificationLevel): string {
  const title = escapeHtml(level.bannerTitle || `Classification: ${level.code}`);
  const text = escapeHtml(level.bannerText);

  return `${OUTLOOK_BANNER_START_MARKER}
<div style="border:1px solid #999;padding:8px;margin-bottom:12px;font-family:Arial,sans-serif;font-size:12px;color:${level.bannerColor};background:${level.bannerBackground};">
  <strong>${title}</strong><br>
  ${text}
</div>
${OUTLOOK_BANNER_END_MARKER}`;
}

function addOrReplaceBanner(currentHtmlBody: string, bannerHtml: string): string {
  const bodyWithoutOldBanner = removeExistingBanner(currentHtmlBody || "");
  const bodyTagMatch = bodyWithoutOldBanner.match(/<body\b[^>]*>/i);

  if (!bodyTagMatch || bodyTagMatch.index === undefined) {
    return `${bannerHtml}${bodyWithoutOldBanner}`;
  }

  const insertIndex = bodyTagMatch.index + bodyTagMatch[0].length;

  return `${bodyWithoutOldBanner.slice(0, insertIndex)}${bannerHtml}${bodyWithoutOldBanner.slice(
    insertIndex
  )}`;
}

function removeExistingBanner(htmlBody: string): string {
  const startIndex = htmlBody.indexOf(OUTLOOK_BANNER_START_MARKER);
  const endIndex = htmlBody.indexOf(OUTLOOK_BANNER_END_MARKER);

  if (startIndex === -1 || endIndex === -1 || endIndex < startIndex) {
    return htmlBody;
  }

  const endOfMarker = endIndex + OUTLOOK_BANNER_END_MARKER.length;

  return `${htmlBody.slice(0, startIndex)}${htmlBody.slice(endOfMarker)}`;
}

async function applySubjectPrefix(
  item: Office.MessageCompose,
  level: ClassificationLevel
): Promise<boolean> {
  const currentSubject = await getSubject(item);
  const subjectWithoutClassifyMePrefix = (currentSubject || "").replace(
    CLASSIFYME_SUBJECT_PREFIX_PATTERN,
    ""
  );
  const updatedSubject = `[${level.code}] ${subjectWithoutClassifyMePrefix}`.trim();

  if (updatedSubject === currentSubject) {
    return false;
  }

  await setSubject(item, updatedSubject);

  return true;
}

async function removeSubjectPrefix(item: Office.MessageCompose): Promise<boolean> {
  const currentSubject = await getSubject(item);
  const updatedSubject = (currentSubject || "").replace(CLASSIFYME_SUBJECT_PREFIX_PATTERN, "");

  if (updatedSubject === currentSubject) {
    return false;
  }

  await setSubject(item, updatedSubject);

  return true;
}

async function trySaveCustomProperties(
  item: Office.MessageCompose,
  level: ClassificationLevel
): Promise<boolean> {
  try {
    const customProperties = await loadCustomProperties(item);
    const updatedAt = new Date().toISOString();

    customProperties.set(CLASSIFICATION_PROPERTY_NAMES.level, level.code);
    customProperties.set(CLASSIFICATION_PROPERTY_NAMES.label, level.label);
    customProperties.set(CLASSIFICATION_PROPERTY_NAMES.updatedAt, updatedAt);
    customProperties.set(CLASSIFICATION_PROPERTY_NAMES.tool, CLASSIFICATION_TOOL_NAME);

    // Outlook custom properties use an async save call and may fail offline or on unsupported accounts.
    await saveCustomProperties(customProperties);

    return true;
  } catch (error) {
    console.info("ClassifyMe n'a pas pu enregistrer les proprietes Outlook.", error);
    return false;
  }
}

function getBodyHtml(item: Office.MessageCompose): Promise<string> {
  return new Promise((resolve, reject) => {
    item.body.getAsync(Office.CoercionType.Html, (result) => {
      handleAsyncResult(result, resolve, reject);
    });
  });
}

function setBodyHtml(item: Office.MessageCompose, htmlBody: string): Promise<void> {
  return new Promise((resolve, reject) => {
    item.body.setAsync(htmlBody, { coercionType: Office.CoercionType.Html }, (result) => {
      handleAsyncResult(result, resolve, reject);
    });
  });
}

function getSubject(item: Office.MessageCompose): Promise<string> {
  return new Promise((resolve, reject) => {
    item.subject.getAsync((result) => {
      handleAsyncResult(result, resolve, reject);
    });
  });
}

function setSubject(item: Office.MessageCompose, subject: string): Promise<void> {
  return new Promise((resolve, reject) => {
    item.subject.setAsync(subject, (result) => {
      handleAsyncResult(result, resolve, reject);
    });
  });
}

function loadCustomProperties(item: Office.MessageCompose): Promise<Office.CustomProperties> {
  return new Promise((resolve, reject) => {
    item.loadCustomPropertiesAsync((result) => {
      handleAsyncResult(result, resolve, reject);
    });
  });
}

function saveCustomProperties(customProperties: Office.CustomProperties): Promise<void> {
  return new Promise((resolve, reject) => {
    customProperties.saveAsync((result) => {
      handleAsyncResult(result, resolve, reject);
    });
  });
}

function handleAsyncResult<T>(
  result: Office.AsyncResult<T>,
  resolve: (value: T) => void,
  reject: (reason?: unknown) => void
): void {
  if (result.status === Office.AsyncResultStatus.Succeeded) {
    resolve(result.value);
    return;
  }

  reject(new Error(result.error?.message || result.error?.code || "L'appel API Outlook a echoue."));
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
