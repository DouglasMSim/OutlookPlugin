// Minimal simulated Office.js for browser smoke tests. Served in place of the
// real CDN script. Top-level pages get a mailbox; iframes get the dialog API.
(function () {
  "use strict";
  var AsyncResultStatus = { Succeeded: "succeeded", Failed: "failed" };
  var EventType = {
    DialogMessageReceived: "dialogMessageReceived",
    DialogEventReceived: "dialogEventReceived",
    DialogParentMessageReceived: "dialogParentMessageReceived",
  };
  var MailboxEnums = {
    ItemNotificationMessageType: {
      ProgressIndicator: "progressIndicator",
      InformationalMessage: "informationalMessage",
      ErrorMessage: "errorMessage",
    },
    ItemType: { Message: "message", Appointment: "appointment" },
    AttachmentType: { File: "file", Item: "item", Cloud: "cloud", Base64: "base64" },
  };
  function onReady(cb) {
    setTimeout(function () {
      if (cb) cb({ host: "Outlook", platform: "OfficeOnline" });
    }, 0);
    return Promise.resolve({ host: "Outlook", platform: "OfficeOnline" });
  }

  var parentMock = window.parent !== window ? window.parent.__officeMock : undefined;
  if (parentMock) {
    // Dialog page inside the simulated dialog iframe.
    window.Office = {
      AsyncResultStatus: AsyncResultStatus,
      EventType: EventType,
      MailboxEnums: MailboxEnums,
      onReady: onReady,
      context: {
        requirements: { isSetSupported: function () { return true; } },
        ui: {
          messageParent: function (message) { parentMock.fromDialog(message); },
          addHandlerAsync: function (type, handler, cb) {
            window.__dialogParentHandler = handler;
            if (cb) cb({ status: AsyncResultStatus.Succeeded });
          },
        },
      },
    };
    return;
  }

  var scenario = window.__scenario || { attachments: [], contents: {} };
  var state = {
    associated: {},
    notifications: [],
    displayCalls: [],
    reads: [],
    completed: 0,
    dialogUrls: [],
    childMessages: [],
    dialogClosed: false,
  };
  window.__officeMock = state;

  var item = {
    itemType: "message",
    itemId: scenario.itemId || "ITEM-1",
    attachments: scenario.attachments,
    getAttachmentContentAsync: function (id, cb) {
      state.reads.push(id);
      setTimeout(function () {
        var c = scenario.contents[id];
        if (!c || c.fail) {
          cb({ status: AsyncResultStatus.Failed, error: { code: 9999, name: "SimulatedFailure", message: "x" } });
        } else {
          cb({ status: AsyncResultStatus.Succeeded, value: { format: c.format || "base64", content: c.content } });
        }
      }, scenario.readDelayMs || 5);
    },
    displayReplyAllFormAsync: function (formData, cb) {
      state.displayCalls.push(formData);
      setTimeout(function () { cb({ status: AsyncResultStatus.Succeeded }); }, 1);
    },
    notificationMessages: {
      replaceAsync: function (key, details, cb) {
        state.notifications.push({ op: "replace", key: key, type: details.type, message: details.message });
        if (cb) cb({ status: AsyncResultStatus.Succeeded });
      },
      removeAsync: function (key, cb) {
        state.notifications.push({ op: "remove", key: key });
        if (cb) cb({ status: AsyncResultStatus.Succeeded });
      },
    },
  };

  window.Office = {
    AsyncResultStatus: AsyncResultStatus,
    EventType: EventType,
    MailboxEnums: MailboxEnums,
    onReady: onReady,
    actions: {
      associate: function (name, fn) { state.associated[name] = fn; },
    },
    context: {
      mailbox: { item: item },
      requirements: {
        isSetSupported: function (name) {
          return name === "Mailbox" ? scenario.supported !== false : true;
        },
      },
      ui: {
        displayDialogAsync: function (url, options, cb) {
          state.dialogUrls.push(url);
          var handlers = {};
          var iframe = document.createElement("iframe");
          iframe.id = "dialog";
          iframe.style.cssText = "position:fixed;inset:0;width:100%;height:100%;border:0;background:#fff";
          iframe.src = url;
          document.body.appendChild(iframe);
          var dialog = {
            addEventHandler: function (type, handler) { handlers[type] = handler; },
            close: function () { iframe.remove(); state.dialogClosed = true; },
            messageChild: function (message) {
              state.childMessages.push(message);
              var h = iframe.contentWindow && iframe.contentWindow.__dialogParentHandler;
              if (h) h({ message: message, origin: window.location.origin });
            },
          };
          state.fromDialog = function (message) {
            var h = handlers[EventType.DialogMessageReceived];
            if (h) h({ message: message, origin: window.location.origin });
          };
          setTimeout(function () { cb({ status: AsyncResultStatus.Succeeded, value: dialog }); }, 0);
        },
      },
    },
  };

  state.invoke = function (name) {
    return new Promise(function (resolve) {
      state.associated[name]({
        completed: function () {
          state.completed += 1;
          resolve();
        },
      });
    });
  };
})();
