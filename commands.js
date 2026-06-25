console.log("[SpecialReplyTools] Script file loaded and executing initial blocks.");

Office.onReady((info) => {
  console.log("[SpecialReplyTools] Office.onReady fired. Host:", info.host, " Platform:", info.platform);
  
  if (info.host === Office.HostType.Outlook) {
    console.log("[SpecialReplyTools] Host is Outlook. Associating 'specialReplyToAll' action...");
    try {
      Office.actions.associate("specialReplyToAll", specialReplyToAll);
      console.log("[SpecialReplyTools] Action successfully associated with the manifest ID.");
    } catch (assocError) {
      console.error("[SpecialReplyTools] Error during action association:", assocError);
    }
  } else {
    console.warn("[SpecialReplyTools] Host type mismatch. Expected Outlook, got:", info.host);
  }
});

function specialReplyToAll(event) {
  console.log("[SpecialReplyTools] ---> specialReplyToAll function triggered! <---");

  const item = Office.context.mailbox.item;
  console.log("[SpecialReplyTools] Context checking: Current Item:", item);

  if (!item) {
    console.error("[SpecialReplyTools] Failure: Office.context.mailbox.item is null or undefined.");
    if (event) {
      console.log("[SpecialReplyTools] Signaling event completion due to missing item context.");
      event.completed();
    }
    return;
  }

  try {
    const currentUserEmail = Office.context.mailbox.userProfile.emailAddress;
    console.log("[SpecialReplyTools] Current User Email discovered:", currentUserEmail);

    let newToRecipients = [];
    let newCcRecipients = [];

    // 1. Map original sender to TO
    if (item.from) {
      console.log("[SpecialReplyTools] Found sender (From):", item.from.emailAddress);
      newToRecipients.push({ displayName: item.from.displayName, emailAddress: item.from.emailAddress });
    } else {
      console.warn("[SpecialReplyTools] Item 'from' attribute is missing.");
    }

    // 2. Map original TOs to CC (excluding self)
    console.log("[SpecialReplyTools] Processing original TO list. Size:", item.to ? item.to.length : 0);
    if (item.to) {
      item.to.forEach((rcp) => {
        if (rcp.emailAddress.toLowerCase() !== currentUserEmail.toLowerCase()) {
          console.log("[SpecialReplyTools] Adding original TO to new CC list:", rcp.emailAddress);
          newCcRecipients.push({ displayName: rcp.displayName, emailAddress: rcp.emailAddress });
        } else {
          console.log("[SpecialReplyTools] Filtering out current user from TO -> CC conversion:", rcp.emailAddress);
        }
      });
    }

    // 3. Keep original CCs in CC (excluding self)
    console.log("[SpecialReplyTools] Processing original CC list. Size:", item.cc ? item.cc.length : 0);
    if (item.cc) {
      item.cc.forEach((rcp) => {
        if (rcp.emailAddress.toLowerCase() !== currentUserEmail.toLowerCase()) {
          console.log("[SpecialReplyTools] Keeping original CC on new CC list:", rcp.emailAddress);
          newCcRecipients.push({ displayName: rcp.displayName, emailAddress: rcp.emailAddress });
        } else {
          console.log("[SpecialReplyTools] Filtering out current user from CC tracking:", rcp.emailAddress);
        }
      });
    }

    const replySubject = item.subject.toLowerCase().startsWith("re:") ? item.subject : "RE: " + item.subject;
    console.log("[SpecialReplyTools] Prepared Subject line:", replySubject);
    console.log("[SpecialReplyTools] Final Payload Assembly -> TO:", newToRecipients, "CC:", newCcRecipients);

    console.log("[SpecialReplyTools] Dispatching displayNewMessageFormAsync call...");
    Office.context.mailbox.displayNewMessageFormAsync(
      {
        toRecipients: newToRecipients,
        ccRecipients: newCcRecipients,
        subject: replySubject,
        htmlBody: "<br><br>"
      },
      function (asyncResult) {
        console.log("[SpecialReplyTools] displayNewMessageFormAsync callback triggered. Status:", asyncResult.status);
        
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error("[SpecialReplyTools] Async window creation error payload:", asyncResult.error);
        } else {
          console.log("[SpecialReplyTools] New message window rendered successfully.");
        }

        if (event) {
          console.log("[SpecialReplyTools] Signaling final event completion to Outlook.");
          event.completed();
        }
      }
    );

  } catch (runtimeError) {
    console.error("[SpecialReplyTools] General runtime crash during execution:", runtimeError);
    if (event) {
      console.log("[SpecialReplyTools] Forcing execution completion handle during exception crash.");
      event.completed();
    }
  }
}