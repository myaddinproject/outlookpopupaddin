Office.actions.associate("onMessageSendHandler", onMessageSendHandler);

function onMessageSendHandler(event) {
    const item = Office.context.mailbox.item;

    // Skip for meeting invites
    if (item.itemType !== Office.MailboxEnums.ItemType.Message) {
        event.completed({ allowEvent: true });
        return;
    }

    // Show popup
    Office.context.ui.displayDialogAsync(
        window.location.origin + "/popup.html",
        { height: 30, width: 20, displayInIframe: true },
        function (asyncResult) {
            const dialog = asyncResult.value;
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, arg => {
                if (arg.message === "yes") {
                    event.completed({ allowEvent: true });
                } else {
                    event.completed({ allowEvent: false });
                }
                dialog.close();
            });
        }
    );
}
