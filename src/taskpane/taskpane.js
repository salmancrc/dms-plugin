/* global document, Office */
import config from "../config";

let insertAt = document.getElementById("text-item");
let dmsButton = document.getElementById("dms-button");
let loaderIcon = document.getElementById("loader-area");

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    syncWithDMS();
    document.getElementById("app-body").style.display = "flex";
  }
});

const insertAtStatus = (value) => {
  insertAt.style.display = value;
}

const dmsBtnStatus = (value) => {
  dmsButton.style.display = value;
}

const loaderStatus = (value) => {
  loaderIcon.style.display = value;
}

const setHTMLText = (text) => {
  insertAt.innerHTML = text;
}

export async function syncWithDMS() {
  const item = Office.context.mailbox.item;

  const emailData = {
    subject: item.subject,
    to: item.to,
    from: item.from,
    body: 'N/A',
    attachments: await getAttachmentsContent(item),
    dateTimeCreated: item.dateTimeCreated
  }

  item.body.getAsync(Office.CoercionType.Text, function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      emailData.body = result.value || 'N/A';
    }

    if (emailData) {
      displayEmailData(emailData);
      loaderStatus("none");
      dmsBtnStatus("block");

      dmsButton.onclick = () => sendEmail(emailData);
    }
  });
}

function displayEmailData(data) {
  insertAtStatus("block");

  const addTextWithLineBreak = (label, value, isAttachment = false) => {
    const displayValue = value || "N/A";
    if (isAttachment) {
      return `
        <dl class="email-attachments">
          <dt>${label}</dt>
          <dd>${displayValue.split(", ").map(name =>
        `<span class="attachment-item">${name}</span>`).join(" ")}
          </dd>
        </dl>`;
    }
    return `<dl class="email-details">
      <dt>${label}</dt>
      <dd>${displayValue}</dd>
    </dl>`;
  };

  const emailDetails = [
    { label: "From: ", value: data?.from?.displayName },
    { label: "To: ", value: data?.to?.map(recipient => recipient.emailAddress || "N/A").join(", ") },
    { label: "Subject: ", value: data?.subject },
    {
      label: "Created On: ",
      value: data?.dateTimeCreated
        ? new Date(data.dateTimeCreated).toLocaleString('en-UK', {
          year: 'numeric', month: 'long', day: 'numeric',
          hour: 'numeric', minute: 'numeric', hour12: true
        })
        : "N/A"
    },
    { label: "Attachments: ", value: data?.attachments?.map(file => file.name || "N/A").join(", "), isAttachment: true },
  ];

  const emailDetailsHtml = emailDetails.map(detail =>
    addTextWithLineBreak(detail.label, detail.value, detail.isAttachment)).join("");

  setHTMLText(emailDetailsHtml);
}

async function getAttachmentsContent(item) {
  if (!item.attachments || item.attachments.length === 0) return [];

  const attachments = [];
  for (const attachment of item.attachments) {
    if (attachment.attachmentType === 'file') {
      const content = await getFileContent(attachment.id);
      attachments.push({
        id: attachment.id,
        name: attachment.name,
        contentType: attachment.contentType,
        size: attachment.size,
        attachmentType: attachment.attachmentType,
        content: content
      });
    }
  }
  return attachments;
}

function getFileContent(attachmentId) {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox.item;
    item.getAttachmentContentAsync(attachmentId, result => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value.content);
      } else {
        reject(new Error(`Error getting attachment content: ${result.error.message}`));
      }
    });
  });
}

async function sendEmail(emailData) {
  loaderStatus("block");
  insertAtStatus("none");

  try {
    const response = await fetch(`${config.apiBaseUrl}/upload-document/`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(emailData),
    });

    loaderStatus("none");
    insertAtStatus("block");
    const data = await response.json();

    if (response.ok) {
      dmsBtnStatus("none");
      setHTMLText(`<span style="color: green;">${data.message}</span>`);
    } else {
      handleError(response.status, response.statusText, data?.error);
    }
  } catch (error) {
    loaderStatus("none");
    insertAtStatus("block");
    setHTMLText(`<span style="color: red;">${error.message}</span>`);
    console.error("Error:::: ", error);
  }
}

function handleError(status, statusText, errorMessage = "") {
  let errorMsg = `Error ${status}: ${statusText}`;

  switch (status) {
    case 404:
      errorMsg = errorMessage || "Resource not found";
      break;
    case 401:
      errorMsg = "Unauthorized - Please check your authentication.";
      break;
    case 500:
      errorMsg = "Internal Server Error";
      break;
    default:
      if (errorMessage) errorMsg = errorMessage;
      break;
  }

  setHTMLText(`<span style="color: red;">${errorMsg}</span>`);
}