Office.actions.associate("onMessageSendHandler", onMessageSendHandler);

function onMessageSendHandler(event) {
    const item = Office.context.mailbox.item;

    // Skip for meeting invites and other non-mail items
    if (item.itemType !== Office.MailboxEnums.ItemType.Message) {
        event.completed({ allowEvent: true });
        return;
    }

    // Show confirmation popup
    Office.context.ui.displayDialogAsync(
        "http://myaddinproject.github.io/outlookpopupaddin/popup.html",
        { height: 30, width: 20, displayInIframe: true },
        function (asyncResult) {
            const dialog = asyncResult.value;
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
                if (arg.message === "yes") {
                    event.completed({ allowEvent: true });  // Send the email
                } else {
                    event.completed({ allowEvent: false }); // Cancel send, keep in drafts
                }
                dialog.close();
            });
        }
    );
}

