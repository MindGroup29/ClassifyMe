/*
 * Outlook-specific classification behavior.
 * The MVP only supports message compose mode: it updates the draft body and, optionally, the subject.
 */

/* global Office, DOMParser, NodeFilter, console, Document, Element, Comment, ChildNode */

import {
  CLASSIFICATION_LEVELS,
  CLASSIFICATION_PROPERTY_NAMES,
  CLASSIFICATION_TOOL_NAME,
  ClassificationLevel,
  OUTLOOK_BANNER_ELEMENT_ID,
  OUTLOOK_BANNER_END_MARKER,
  OUTLOOK_BANNER_START_MARKER,
  OUTLOOK_WEB_HTML_ID_PREFIX,
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

  return `<div id="${OUTLOOK_BANNER_ELEMENT_ID}" style="border:1px solid #999;padding:8px;margin-bottom:12px;font-family:Arial,sans-serif;font-size:12px;color:${level.bannerColor};background:${level.bannerBackground};">
  <strong>${title}</strong><br>
  ${text}
</div>`;
}

function addOrReplaceBanner(currentHtmlBody: string, bannerHtml: string): string {
  if (typeof DOMParser === "undefined") {
    throw new Error(
      "Le client Outlook ne permet pas de mettre a jour le bandeau de classification."
    );
  }

  const hasHtmlDocument = /<html\b/i.test(currentHtmlBody || "");
  const parsedDocument = new DOMParser().parseFromString(currentHtmlBody || "", "text/html");

  /*
   * The pilot banner used HTML comments as its only identifier. Outlook on the web and
   * new Outlook can rewrite the body and discard those comments, so the previous lookup
   * could no longer find the existing banner. The new banner has a real div id instead.
   */
  removeExistingBanners(parsedDocument);
  parsedDocument.body.insertAdjacentHTML("afterbegin", bannerHtml);

  // Preserve the response HTML as a fragment when Outlook supplied a fragment, and as a
  // complete document when it supplied one. Only ClassifyMe banner nodes are changed.
  return hasHtmlDocument ? parsedDocument.documentElement.outerHTML : parsedDocument.body.innerHTML;
}

function removeExistingBanners(htmlDocument: Document): void {
  htmlDocument.querySelectorAll("[id]").forEach((element) => {
    if (isClassifyMeBannerElement(element)) {
      element.remove();
    }
  });

  removeLegacyCommentBanners(htmlDocument);
  removeRecognizableLegacyBanners(htmlDocument);
}

function isClassifyMeBannerElement(element: Element): boolean {
  /*
   * Outlook on the web can return an id as x_classifyme-classification-banner.
   * It can add the prefix again to quoted or forwarded HTML, so remove every leading
   * x_ before comparing with the identifier written by ClassifyMe.
   */
  let normalizedId = element.id;

  while (normalizedId.startsWith(OUTLOOK_WEB_HTML_ID_PREFIX)) {
    normalizedId = normalizedId.slice(OUTLOOK_WEB_HTML_ID_PREFIX.length);
  }

  return normalizedId === OUTLOOK_BANNER_ELEMENT_ID;
}

function removeLegacyCommentBanners(htmlDocument: Document): void {
  const comments: Comment[] = [];
  const commentWalker = htmlDocument.createTreeWalker(htmlDocument, NodeFilter.SHOW_COMMENT);
  let currentNode = commentWalker.nextNode();

  while (currentNode) {
    comments.push(currentNode as Comment);
    currentNode = commentWalker.nextNode();
  }

  let startComment: Comment | undefined;
  comments.forEach((comment) => {
    if (isLegacyMarkerComment(comment, OUTLOOK_BANNER_START_MARKER)) {
      startComment = comment;
      return;
    }

    if (startComment && isLegacyMarkerComment(comment, OUTLOOK_BANNER_END_MARKER)) {
      removeSiblingRange(startComment, comment);
      startComment = undefined;
    }
  });
}

function isLegacyMarkerComment(comment: Comment, markerHtml: string): boolean {
  return `<!-- ${comment.data.trim()} -->` === markerHtml;
}

function removeSiblingRange(startNode: ChildNode, endNode: ChildNode): void {
  // The old generated markup placed both marker comments next to the banner in one parent.
  // Do not attempt to remove a malformed range across parents: preserving user content wins.
  if (!startNode.parentNode || startNode.parentNode !== endNode.parentNode) {
    return;
  }

  let nodeToRemove: ChildNode | null = startNode;
  while (nodeToRemove) {
    const nextNode: ChildNode | null = nodeToRemove.nextSibling;
    nodeToRemove.remove();

    if (nodeToRemove === endNode) {
      return;
    }

    nodeToRemove = nextNode;
  }
}

function removeRecognizableLegacyBanners(htmlDocument: Document): void {
  htmlDocument.querySelectorAll("div").forEach((element) => {
    if (isRecognizableLegacyBanner(element)) {
      element.remove();
    }
  });
}

function isRecognizableLegacyBanner(element: Element): boolean {
  const style = element.getAttribute("style") || "";
  const title = element.querySelector("strong, b")?.textContent || "";
  const visibleText = normalizeHtmlText(element.textContent || "");

  // This fallback is only for drafts produced by the pilot before the id existed. It requires
  // the complete generated title, message and characteristic styling; the title alone is never
  // used to identify a banner. Outlook Web can expand border: into border-width/style/color.
  return (
    hasLegacyBannerStyle(style) &&
    CLASSIFICATION_LEVELS.some((level) => {
      const expectedTitle = level.bannerTitle || `Classification: ${level.code}`;
      const expectedText = normalizeHtmlText(`${expectedTitle} ${level.bannerText}`);

      return (
        normalizeHtmlText(title) === normalizeHtmlText(expectedTitle) &&
        visibleText === expectedText
      );
    })
  );
}

function hasLegacyBannerStyle(style: string): boolean {
  return (
    /(?:^|;)\s*border(?:-width)?\s*:/i.test(style) &&
    /(?:^|;)\s*margin-bottom\s*:/i.test(style) &&
    /(?:^|;)\s*padding\s*:/i.test(style)
  );
}

function normalizeHtmlText(value: string): string {
  return value.replace(/\s+/g, " ").trim();
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
