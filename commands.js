// Register the function with the Office environment
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    // This associates the string name in the manifest with the actual JS function
    Office.actions.associate("specialReplyToAll", specialReplyToAll);
  }
});

function specialReplyToAll(event) {
  const debugUrl = "https://engine.rewst.io/webhooks/custom/trigger/0198d29a-a418-71e5-a3bc-fa06f26b1b5d/0191e80b-fecf-7c05-8234-156ea48fc2cb";
  
  const debugPayload = {
    timestamp: new Date().toISOString(),
    message: "specialReplyToAll function triggered successfully"
  };

  fetch(debugUrl, {
    method: "POST",
    node: "no-cors",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify(debugPayload)
  })
  .then(response => console.log("Debug log sent successfully"))
  .catch(error => console.error("Failed to send debug log:", error));
  // --------------------------------

  const item = Office.context.mailbox.item;
  if (!item) {
    if (event) event.completed();
    return;
  }

  const currentUserEmail = Office.context.mailbox.userProfile.emailAddress;
  let newToRecipients = [];
  let newCcRecipients = [];

  if (item.from) {
    newToRecipients.push({ displayName: item.from.displayName, emailAddress: item.from.emailAddress });
  }

  item.to.forEach((rcp) => {
    if (rcp.emailAddress.toLowerCase() !== currentUserEmail.toLowerCase()) {
      newCcRecipients.push({ displayName: rcp.displayName, emailAddress: rcp.emailAddress });
    }
  });

  item.cc.forEach((rcp) => {
    if (rcp.emailAddress.toLowerCase() !== currentUserEmail.toLowerCase()) {
      newCcRecipients.push({ displayName: rcp.displayName, emailAddress: rcp.emailAddress });
    }
  });

  const replySubject = item.subject.toLowerCase().startsWith("re:") ? item.subject : "RE: " + item.subject;

  Office.context.mailbox.displayNewMessageFormAsync(
    {
      toRecipients: newToRecipients,
      ccRecipients: newCcRecipients,
      subject: replySubject,
      htmlBody: "<br><br>"
    },
    function (asyncResult) {
      if (event) event.completed();
    }
  );
}