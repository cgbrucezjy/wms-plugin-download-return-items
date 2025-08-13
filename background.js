// background.js

/**
 * Listens for messages from the content script. The primary purpose is to act as a proxy
 * for fetching cross-origin images, which the content script cannot do by itself due to CORS.
 */
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  // Check if the message is a request to fetch an image
  if (request.action === "fetchImageAsBase64") {
    // Use the fetch API to get the image data.
    // This works because the background script has host_permissions from the manifest.
    fetch(request.url)
      .then(response => {
        if (!response.ok) {
          throw new Error(`Network response was not ok: ${response.statusText}`);
        }
        return response.blob(); // Get the image as a Blob object
      })
      .then(blob => {
        // Use a FileReader to convert the Blob into a Base64 data URL
        const reader = new FileReader();
        reader.onloadend = () => {
          // Once conversion is complete, send the successful response back to the content script
          sendResponse({ success: true, base64: reader.result });
        };
        reader.onerror = () => {
          sendResponse({ success: false, error: 'Failed to read blob as Base64' });
        };
        reader.readAsDataURL(blob);
      })
      .catch(error => {
        // If any error occurs during the fetch, send an error response back
        console.error('Background fetch error:', error);
        sendResponse({ success: false, error: error.message });
      });

    // **Important**: Return true to indicate that sendResponse will be called asynchronously.
    return true;
  }
});