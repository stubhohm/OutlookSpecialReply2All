console.log("[SpecialReplyTools] Script file loaded and executing initial blocks.");

window.specialReplyToAll = function (event) {
  console.log("[SpecialReplyTools] ---> specialReplyToAll function triggered! <---");

  const item = Office.context.mailbox.item;
  if (!item) {
    console.error("[SpecialReplyTools] Context checking failed: Item is null.");
    if (event) event.completed();
    return;
  }

  try {
    const currentUserEmail = Office.context.mailbox.userProfile.emailAddress;
    let newToRecipients = [];
    let newCcRecipients = [];

    if (item.from) {
      newToRecipients.push({ displayName: item.from.displayName, emailAddress: item.from.emailAddress });
    }

    if (item.to) {
      item.to.forEach((rcp) => {
        if (rcp.emailAddress.toLowerCase() !== currentUserEmail.toLowerCase()) {
          newCcRecipients.push({ displayName: rcp.displayName, emailAddress: rcp.emailAddress });
        }
      });
    }

    if (item.cc) {
      item.cc.forEach((rcp) => {
        if (rcp.emailAddress.toLowerCase() !== currentUserEmail.toLowerCase()) {
          newCcRecipients.push({ displayName: rcp.displayName, emailAddress: rcp.emailAddress });
        }
      });
    }

    const replySubject = item.subject.toLowerCase().startsWith("re:") ? item.subject : "RE: " + item.subject;

    Office.context.mailbox.displayNewMessageFormAsync(
      {
        toRecipients: newToRecipients,
        ccRecipients: newCcRecipients,
        subject: replySubject,
        htmlBody: "<br><br>"
      },
      function (asyncResult) {
        console.log("[SpecialReplyTools] Window creation complete. Status:", asyncResult.status);
        // Crucial: Tell Outlook the operation has finished
        if (event) event.completed();
      }
    );

  } catch (runtimeError) {
    console.error("[SpecialReplyTools] Crash during execution:", runtimeError);
    if (event) event.completed();
  }
};


Office.onReady((info) => {
  console.log("[SpecialReplyTools] Office.onReady fired. Host:", info.host);
  if (info.host === Office.HostType.Outlook) {
    try {
      Office.actions.associate("specialReplyToAll", window.specialReplyToAll);
      console.log("[SpecialReplyTools] Action successfully associated.");
    } catch (assocError) {
      console.error("[SpecialReplyTools] Error during action association:", assocError);
    }
  }
});
