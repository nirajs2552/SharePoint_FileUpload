
/**
 * Launches SharePoint/OneDrive File Picker v8 as a popup via POST.
 * Returns a Promise that resolves with the selected items.
 */
export function launchFilePicker({ baseUrl, accessToken }) {
  return new Promise((resolve, reject) => {
    const pickerWindow = window.open("", "MSPicker", "width=1080,height=720");
    if (!pickerWindow) return reject(new Error("Popup blocked"));
    const channelId = crypto.randomUUID();
    const origin = window.location.origin;

    // Listen for the picker postMessage
    function onMessage(event) {
      // In production validate event.origin matches baseUrl
      const data = event.data;
      if (data && data.type === "picker/close") {
        window.removeEventListener("message", onMessage);
        try { pickerWindow.close(); } catch {}
        if (data.result && data.result.items) {
          resolve(data.result.items);
        } else {
          resolve([]);
        }
      }
    }
    window.addEventListener("message", onMessage);

    // Build the picker v8 options
    const options = {
      sdk: "8.0",
      entry: { sharePoint: {} }, // or { oneDrive: {} }
      authentication: {},
      messaging: { origin, channelId },
      selection: { mode: "files", allowMultiple: true }
    };

    const url = `${baseUrl}/_layouts/15/FilePicker.aspx`;
    const formHtml = `
      <form method="POST" action="${url}">
        <input type="hidden" name="access_token" value="${accessToken}">
        <input type="hidden" name="filePicker" value='${JSON.stringify(options)}'>
      </form>
      <script>document.forms[0].submit();</script>
    `;
    pickerWindow.document.write(formHtml);
    pickerWindow.document.close();
  });
}
