/* global Office */
(function () {
  "use strict";
  var address = "sync-639acf3e3a5d76cfe1a236bc199dec9d85b73184@inbound.crmbook.dk";
  function addCrmBcc(event) {
    var item = Office.context.mailbox.item;
    var completed = false;
    function finish(message, failed) {
      function done() { if (!completed) { completed = true; event.completed(); } }
      try {
        item.notificationMessages.replaceAsync("crmbook-bcc", failed ? {
          type: "errorMessage", message: message
        } : {
          type: "informationalMessage", message: message, icon: "Icon.16", persistent: false
        }, done);
      } catch (error) { done(); }
    }
    try {
      if (!item || !item.bcc) { finish("Open a message for editing first.", true); return; }
      item.bcc.getAsync(function (result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          finish("Could not read BCC. Please try again.", true); return;
        }
        var present = result.value.some(function (recipient) {
          return (recipient.emailAddress || "").toLowerCase() === address;
        });
        if (present) { finish("CRMbook is already in BCC.", false); return; }
        item.bcc.addAsync([address], function (added) {
          finish(added.status === Office.AsyncResultStatus.Succeeded ?
            "CRMbook added to BCC." : "Could not add CRMbook. Please check BCC and try again.",
            added.status !== Office.AsyncResultStatus.Succeeded);
        });
      });
    } catch (error) { finish("Could not add CRMbook. Please try again.", true); }
  }
  Office.onReady(function () { Office.actions.associate("addCrmBcc", addCrmBcc); });
}());
