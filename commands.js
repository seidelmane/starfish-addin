/* global Office */
const STARFISH_BASE =
  "https://cod.starfishsolutions.com/starfish-ops/student/students.html?tabRequest=allStudentsTab#studentList";

Office.onReady(() => {});

function getFromAddress() {
  const item = Office.context.mailbox.item;
  return item?.from?.emailAddress?.trim() || "";
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
      event.completed(); return;
    }
    const starfishUrl = `${STARFISH_BASE}&email=${encodeURIComponent(from)}`;
    const relayUrl = `https://seidelmane.github.io/starfish-addin/relay.html#${encodeURIComponent(starfishUrl)}`;
    Office.context.ui.displayDialogAsync(relayUrl, { height: 45, width: 30, requireHTTPS: true }, () => event.completed());
  } catch { event.completed(); }
}

Office.actions.associate("openStarfish", openStarfish);
