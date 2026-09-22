/*
 * Rappel de classification déclenché lors d'une nouvelle composition Outlook.
 * Le runtime reste volontairement autonome : il affiche un rappel non bloquant,
 * puis laisse le task pane existant gérer la classification après le clic utilisateur.
 */

/* global Office, console */

// Outlook limite les clés de notification à 32 caractères.
const CLASSIFICATION_REMINDER_KEY = "classifyme-reminder";

type ClassificationReminderEvent = "OnNewMessageCompose" | "OnNewAppointmentOrganizer";

function onNewMessageCompose(event: Office.MailboxEvent): void {
  void showClassificationReminder("OnNewMessageCompose", event);
}

function onNewAppointmentOrganizer(event: Office.MailboxEvent): void {
  void showClassificationReminder("OnNewAppointmentOrganizer", event);
}

async function showClassificationReminder(
  eventName: ClassificationReminderEvent,
  event: Office.MailboxEvent
): Promise<void> {
  try {
    const notificationMessages = Office.context.mailbox?.item?.notificationMessages;

    if (!notificationMessages) {
      return;
    }

    await addReminderNotification(notificationMessages, eventName);
  } catch (error) {
    // Le rappel reste non bloquant, mais l'erreur demeure visible pour le support technique.
    console.error("[ClassifyMe][Rappel Outlook] Impossible d'afficher le rappel.", error);
  } finally {
    completeEvent(event);
  }
}

function addReminderNotification(
  notificationMessages: Office.NotificationMessages,
  eventName: ClassificationReminderEvent
): Promise<void> {
  const isAppointment = eventName === "OnNewAppointmentOrganizer";
  const notificationTitle = isAppointment
    ? "Classification de la réunion"
    : "Classification du message";
  const commandId = isAppointment
    ? "ClassifyMeAppointmentComposeTaskpaneButton"
    : "ClassifyMeComposeTaskpaneButton";

  return new Promise((resolve, reject) => {
    notificationMessages.addAsync(
      CLASSIFICATION_REMINDER_KEY,
      {
        type: Office.MailboxEnums.ItemNotificationMessageType.InsightMessage,
        message: `${notificationTitle} — Pensez à classifier ${
          isAppointment ? "cette réunion" : "ce message"
        } avant son envoi.`,
        icon: "Icon.16x16",
        actions: [
          {
            actionType: "showTaskPane",
            actionText: "Ouvrir ClassifyMe",
            commandId,
            contextData: "classification-reminder",
          },
        ],
      } as Office.NotificationMessageDetails,
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
          return;
        }

        reject(result.error);
      }
    );
  });
}

function completeEvent(event: Office.MailboxEvent): void {
  try {
    event.completed();
  } catch {
    // Outlook doit rester utilisable même si le runtime refuse un second achèvement.
  }
}

// L'association immédiate est requise par l'activation événementielle Outlook.
Office.actions.associate("onNewMessageCompose", onNewMessageCompose);
Office.actions.associate("onNewAppointmentOrganizer", onNewAppointmentOrganizer);
