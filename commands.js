/* global Office */
const STARFISH_BASE =
  "https://cod.starfishsolutions.com/starfish-ops/student/students.html?tabRequest=allStudentsTab#studentList";

Office.onReady(() => { /* no-op */ });

function getFromAddress() {
  const item = Office.context.mailbox.item;
  if (item && item.from && item.from.emailAddress) {
    // This is SMTP in OWA/new Outlook. If EX routing ever appears, add a resolver here.
    return item.from.emailAddress.trim();
  }
  return "";
}

async function openStarfish(event) {
  try {
    const from = getFromAddress();
    if (!from) {
      Office.context.mailbox.item.notificationMessages.replaceAsync("starfish-msg", {
        type: "informationalMessage",
        message: "No From address found on this item.",
        icon: "icon16",
        persistent: false
      });
      event.completed();
      return;
    }

    const starfishUrl = `${STARFISH_BASE}&email=${encodeURIComponent(from)}`;

    // Open via relay to ensure a REAL browser tab (where your Chrome/Edge extension runs)
    const relayUrl = `https://localhost/relay.html#${encodeURIComponent(starfishUrl)}`;
    Office.context.ui.displayDialogAsync(
      relayUrl,
      { height: 45, width: 30, requireHTTPS: true },
      () => event.completed()
    );
  } catch (e) {
    event.completed();
  }
}

Office.actions.associate("openStarfish", openStarfish);
