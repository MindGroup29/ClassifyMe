/*
 * Spike temporaire "spike-autoopen" pour l'activation événementielle Outlook.
 *
 * Outlook ne prend pas en charge les runtimes partagés requis par
 * Office.addin.showAsTaskpane(). Le repli officiellement supporté consiste à
 * afficher une notification Outlook dont l'action ouvre le task pane existant.
 */

/* global Office, console */

const SPIKE_LOG_PREFIX = "[ClassifyMe][Spike AutoOpen]";
const SPIKE_NOTIFICATION_KEY = "classifyme-spike-autoopen";
const SPIKE_DIAGNOSTIC_NOTIFICATION_KEY = "classifyme-spike-diagnostic";

type LaunchEventName = "OnNewMessageCompose" | "OnNewAppointmentOrganizer";

console.log(`${SPIKE_LOG_PREFIX} runtime chargé`);

function onNewMessageCompose(event: Office.MailboxEvent): void {
  handleLaunchEvent("OnNewMessageCompose", event);
}

function onNewAppointmentOrganizer(event: Office.MailboxEvent): void {
  handleLaunchEvent("OnNewAppointmentOrganizer", event);
}

function handleLaunchEvent(eventName: LaunchEventName, event: Office.MailboxEvent): void {
  console.log(`${SPIKE_LOG_PREFIX} ${eventName} déclenché`);
  console.log(`${SPIKE_LOG_PREFIX} handler démarré`);
  console.warn(
    `${SPIKE_LOG_PREFIX} Office.addin.showAsTaskpane() non appelé : ` +
      "l'API exige SharedRuntime 1.1, qui n'est pas pris en charge par Outlook."
  );

  try {
    const mailbox = Office.context.mailbox;
    const item = mailbox && mailbox.item;
    const notificationMessages = item && item.notificationMessages;

    if (!notificationMessages) {
      throw new Error("L'API Outlook notificationMessages n'est pas disponible.");
    }

    console.log(
      `${SPIKE_LOG_PREFIX} ajout de la notification de repli permettant d'ouvrir ClassifyMe`
    );

    notificationMessages.addAsync(
      SPIKE_NOTIFICATION_KEY,
      {
        type: Office.MailboxEnums.ItemNotificationMessageType.InsightMessage,
        message: "ClassifyMe est prêt. Ouvrez le panneau pour classifier cet élément.",
        icon: "Icon.16x16",
        actions: [
          {
            actionType: "showTaskPane",
            actionText: "Ouvrir ClassifyMe",
            commandId: getTaskPaneCommandId(),
            contextData: "spike-autoopen",
          },
        ],
      } as Office.NotificationMessageDetails,
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          console.log(`${SPIKE_LOG_PREFIX} notification de repli ajoutée avec succès`);
          completeLaunchEvent(event, eventName);
        } else {
          console.error(`${SPIKE_LOG_PREFIX} échec de la notification de repli`, result.error);
          addDiagnosticNotification(notificationMessages, event, eventName);
        }
      }
    );
  } catch (error) {
    console.error(`${SPIKE_LOG_PREFIX} échec du handler`, error);
    completeLaunchEvent(event, eventName);
  }
}

function getTaskPaneCommandId(): string {
  const item = Office.context.mailbox.item;

  if (item && item.itemType === Office.MailboxEnums.ItemType.Appointment) {
    return "ClassifyMeAppointmentComposeTaskpaneButton";
  }

  return "ClassifyMeComposeTaskpaneButton";
}

function addDiagnosticNotification(
  notificationMessages: Office.NotificationMessages,
  event: Office.MailboxEvent,
  eventName: LaunchEventName
): void {
  console.warn(`${SPIKE_LOG_PREFIX} tentative de notification diagnostique minimale`);

  notificationMessages.addAsync(
    SPIKE_DIAGNOSTIC_NOTIFICATION_KEY,
    {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: "Le handler ClassifyMe s'est déclenché, mais l'action d'ouverture a échoué.",
      persistent: true,
    },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        console.log(`${SPIKE_LOG_PREFIX} notification diagnostique ajoutée avec succès`);
      } else {
        console.error(`${SPIKE_LOG_PREFIX} échec de la notification diagnostique`, result.error);
      }

      completeLaunchEvent(event, eventName);
    }
  );
}

function completeLaunchEvent(event: Office.MailboxEvent, eventName: LaunchEventName): void {
  try {
    event.completed();
    console.log(`${SPIKE_LOG_PREFIX} event.completed() exécuté pour ${eventName}`);
  } catch (error) {
    console.error(`${SPIKE_LOG_PREFIX} échec de event.completed() pour ${eventName}`, error);
  }
}

Office.actions.associate("onNewMessageCompose", onNewMessageCompose);
Office.actions.associate("onNewAppointmentOrganizer", onNewAppointmentOrganizer);
console.log(`${SPIKE_LOG_PREFIX} handlers associés`);
