// Register the function with the Office environment
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    // This associates the string name in the manifest with the actual JS function
    Office.actions.associate("specialReplyToAll", specialReplyToAll);
  }
});

function specialReplyToAll(event) {
  const item = Office.context.mailbox.item;
  if (!item) {
    if (event) event.completed();
    return;
  }

  const currentUserEmail = Office.context.mailbox.userProfile.emailAddress;
  let newToRecipients = [];
  let newCcRecipients = [];

  // 1. Map original sender to TO
  if (item.from) {
    newToRecipients.push({ displayName: item.from.displayName, emailAddress: item.from.emailAddress });
  }

  // 2. Map original TOs to CC (excluding self)
  item.to.forEach((rcp) => {
    if (rcp.emailAddress.toLowerCase() !== currentUserEmail.toLowerCase()) {
      newCcRecipients.push({ displayName: rcp.displayName, emailAddress: rcp.emailAddress });
    }
  });

  // 3. Keep original CCs in CC (excluding self)
  item.cc.forEach((rcp) => {
    if (rcp.emailAddress.toLowerCase() !== currentUserEmail.toLowerCase()) {
      newCcRecipients.push({ displayName: rcp.displayName, emailAddress: rcp.emailAddress });
    }
  });

  const replySubject = item.subject.toLowerCase().startsWith("re:") ? item.subject : "RE: " + item.subject;

  // 4. Open the new window layout
  Office.context.mailbox.displayNewMessageFormAsync(
    {
      toRecipients: newToRecipients,
      ccRecipients: newCcRecipients,
      subject: replySubject,
      htmlBody: "<br><br>"
    },
    function (asyncResult) {
      // 5. Signal to Outlook that the execution is finished
      if (event) event.completed();
    }
  );
}