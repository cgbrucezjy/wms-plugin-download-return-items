// content.js

(function () {
  /**
   * Fetches an image as a Base64 string by delegating the request
   * to the extension's background script, which can bypass CORS.
   * It also resizes the image to keep the Excel file size reasonable.
   * @param {string} url The URL of the image to fetch.
   * @returns {Promise<string>} A promise that resolves with the resized Base64 data URL.
   */
  async function imageToBase64(url) {
    return new Promise((resolve, reject) => {
      // Send a message to the background script to fetch and convert the image
      chrome.runtime.sendMessage(
        {
          action: "fetchImageAsBase64",
          url: url
        },
        (response) => {
          if (chrome.runtime.lastError) {
            return reject(new Error(chrome.runtime.lastError.message));
          }

          if (response && response.success) {
            // Once we have the full-size base64, load it into an image
            // element so we can draw it on a canvas for resizing.
            const img = new Image();
            img.onload = function() {
                const canvas = document.createElement('canvas');
                const ctx = canvas.getContext('2d');
                
                const maxWidth = 400; // Max width for the image in Excel
                const maxHeight = 300; // Max height
                let { width, height } = img;
                
                // Calculate new dimensions while maintaining aspect ratio
                if (width > maxWidth) {
                    height = (height * maxWidth) / width;
                    width = maxWidth;
                }
                if (height > maxHeight) {
                    width = (width * maxHeight) / height;
                    height = maxHeight;
                }
                
                canvas.width = width;
                canvas.height = height;
                ctx.drawImage(img, 0, 0, width, height);
                
                // Re-encode to JPEG with a specific quality to reduce file size
                const resizedBase64 = canvas.toDataURL('image/jpeg', 1);
                resolve(resizedBase64);
            };
            img.onerror = () => reject(new Error('Failed to load image from base64 source.'));
            img.src = response.base64; // The base64 data from the background script
          } else {
            const errorMessage = response ? response.error : 'Unknown error from background script';
            reject(new Error(`Failed to convert image: ${errorMessage}`));
          }
        }
      );
    });
  }

  /**
   * Finds all images within the popup iframe, converts them to Base64,
   * and returns an array of image data.
   * @param {string} claimId The ID of the claim, used for naming images.
   * @returns {Promise<Array<object>>} A promise that resolves with an array of image data objects.
   */
  async function handlePopupImages(claimId) {
    console.log(`Processing images for claim ID: ${claimId}`);
    try {
      const topDoc = window.top.document;
      const popupIframe = topDoc.querySelector('iframe.special-orders-show-dialog');
      if (!popupIframe) {
        console.warn("Could not find the popup iframe.");
        return [];
      }
      
      const iframeDoc = popupIframe.contentDocument || popupIframe.contentWindow.document;
      const images = iframeDoc.querySelectorAll('#module-container #image img');
      console.log(`Found ${images.length} images.`);
      
      if (images.length === 0) return [];
      
      const imageDataPromises = [];
      for (let i = 0; i < images.length; i++) {
        const img = images[i];
        const imgSrc = img.src;
        
        if (imgSrc) {
          console.log(`Converting image ${i + 1}: ${imgSrc}`);
          const promise = imageToBase64(imgSrc)
            .then(base64 => ({
              name: `${claimId}_image_${i + 1}`,
              base64: base64,
              src: imgSrc
            }))
            .catch(error => {
              console.warn(`Failed to convert image ${imgSrc}:`, error);
              return {
                name: `${claimId}_image_${i + 1}`,
                base64: null,
                src: imgSrc,
                error: error.message
              };
            });
          imageDataPromises.push(promise);
        }
      }
      
      return Promise.all(imageDataPromises);
    } catch (error) {
      console.error(`Error while handling popup images for ${claimId}:`, error);
      return [];
    }
  }

/**
 * Closes the popup dialog using multiple, increasingly forceful methods
 * to ensure it closes reliably every time.
 */
async function closePopup() {
  console.log("Attempting to close popup dialog...");
  try {
    const topDoc = window.top.document;

    // --- Method 1: Simulate pressing the 'Escape' key (Often the best method) ---
    console.log("Attempt 1: Simulating 'Escape' key press.");
    topDoc.body.dispatchEvent(new KeyboardEvent('keydown', {
      key: 'Escape',
      keyCode: 27,
      bubbles: true,
      cancelable: true,
      view: window.top
    }));
    await new Promise(resolve => setTimeout(resolve, 500)); 

    // Check if it worked
    if (!topDoc.querySelector('.ui-dialog')) {
      console.log("Success: Popup closed with 'Escape' key.");
      return;
    }

    // --- Method 2: Simulate a full, realistic mouse click on the button ---
    const closeButton = topDoc.querySelector('.ui-dialog .ui-dialog-titlebar-close');
    if (closeButton) {
      console.log("Attempt 2: Simulating a full mouse click on the close button.");
      // This is more robust than a simple .click() as it triggers all mouse events.
      const downEvent = new MouseEvent('mousedown', { bubbles: true });
      const upEvent = new MouseEvent('mouseup', { bubbles: true });
      const clickEvent = new MouseEvent('click', { bubbles: true });
      
      closeButton.dispatchEvent(downEvent);
      closeButton.dispatchEvent(upEvent);
      closeButton.dispatchEvent(clickEvent);
      await new Promise(resolve => setTimeout(resolve, 500));
    }

    // Check if it worked
    if (!topDoc.querySelector('.ui-dialog')) {
      console.log("Success: Popup closed by simulating a full click.");
      return;
    }

    // --- Method 3: Brute-force removal (The final fallback) ---
    console.log("Attempt 3 (Fallback): Forcibly removing dialog elements from the DOM.");
    topDoc.querySelectorAll('.ui-dialog, .ui-widget-overlay').forEach(el => el.remove());
    console.log("Cleanup: Forcibly removed dialog and overlay elements.");

  } catch (error) {
    console.error("An error occurred while trying to close the popup:", error);
  }
}

  /**
   * Orchestrates the process for a single row: clicks the button,
   * processes the images from the resulting popup, and closes it.
   * @param {HTMLElement} row The table row element.
   * @param {string} claimId The ID of the claim for this row.
   * @returns {Promise<Array<object>>} A promise that resolves with the image data.
   */
  async function processRowImages(row, claimId) {
    const viewButton = row.querySelector('input.claim'); // This selector should find the button
    if (!viewButton) {
      console.log(`No view button found for claim ID ${claimId}`);
      return [];
    }
    
    console.log(`Clicking view button for claim ID ${claimId}`);
    viewButton.click();
    
    // Wait for the popup iframe to appear and load
    await new Promise(resolve => setTimeout(resolve, 2000));
    
    const imageData = await handlePopupImages(claimId);
    
    await closePopup();
    
    return imageData;
  }

/**
 * Creates a professionally formatted Excel workbook.
 * This version includes a two-level, colored/merged header and 5 image columns.
 * @param {Array<Array<string>>} data The scraped text data for the Excel sheet.
 * @param {object} imageDataMap A map of claimId to its image data array.
 * @returns {Promise<ExcelJS.Workbook>} A promise that resolves with the generated workbook.
 */
async function createExcelWorkbook(data, imageDataMap = {}) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Package Data');

  // --- 1. MANUAL HEADER CREATION ---
  
  // Define the headers from your text
  const mainHeaders = [
    '认领单号', '仓库', '退件跟踪号TRACKING NO.', '包裹数量', '客户ID',
    '有效时间', '认领时间', '创建时间', '完成时间', '更新时间', '仓库收货反馈'
  ];
  const overallImageHeader = '图片';
  const imageSubHeaders = ['图片 1', '图片 2', '图片 3', '图片 4', '图片 5'];

  // Manually set the values for the two header rows
  const headerRow1 = worksheet.getRow(1);
  const headerRow2 = worksheet.getRow(2);
  headerRow1.height = 20;
  headerRow2.height = 20;

  // Populate Row 1
  mainHeaders.forEach((header, index) => {
    worksheet.getCell(1, index + 1).value = header;
  });
  worksheet.getCell(1, 12).value = overallImageHeader;

  // Populate Row 2 with image sub-headers
  imageSubHeaders.forEach((header, index) => {
    worksheet.getCell(2, 12 + index).value = header;
  });

  // Perform cell merges
  // Vertically merge the main headers (A1:A2, B1:B2, etc.)
  for (let i = 1; i <= mainHeaders.length; i++) {
    worksheet.mergeCells(1, i, 2, i);
  }
  // Horizontally merge the overall '图片' header (L1:P1)
  worksheet.mergeCells(1, 12, 1, 16);

  // --- 2. APPLY STYLES TO HEADERS ---
  
  // Iterate through all cells in the first two rows to apply styles
  for (let i = 1; i <= 2; i++) {
    worksheet.getRow(i).eachCell({ includeEmpty: true }, (cell) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFDCE6F1' } // A professional light blue color
      };
      cell.font = {
        bold: true,
        name: 'Calibri',
        size: 11
      };
      cell.alignment = {
        vertical: 'middle',
        horizontal: 'center'
      };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
    });
  }

  // --- 3. DEFINE COLUMN KEYS AND WIDTHS ---
  
  // This step is still needed to map your data keys to the correct columns when using addRow.
  worksheet.columns = [
    { key: 'claimId', width: 15 },
    { key: 'warehouse', width: 12 },
    { key: 'tracking', width: 30 },
    { key: 'qty', width: 10 },
    { key: 'ref', width: 20 },
    { key: 'validTime', width: 22 },
    { key: 'claimTime', width: 22 },
    { key: 'created', width: 22 },
    { key: 'completed', width: 22 },
    { key: 'updated', width: 22 },
    { key: 'notes', width: 40 },
    // Keys for the 5 image columns
    { key: 'image1', width: 22 },
    { key: 'image2', width: 22 },
    { key: 'image3', width: 22 },
    { key: 'image4', width: 22 },
    { key: 'image5', width: 22 }
  ];

  // --- 4. POPULATE DATA ROWS AND IMAGES ---
  
  for (let i = 1; i < data.length; i++) {
    const rowData = data[i];
    const claimId = rowData[0];
    
    // addRow will now correctly place data starting from row 3
    const excelRow = worksheet.addRow({
      claimId: claimId, warehouse: rowData[1], tracking: rowData[2], qty: rowData[3],
      ref: rowData[4], validTime: rowData[5], claimTime: rowData[6], created: rowData[7],
      completed: rowData[8], updated: rowData[9], notes: rowData[10],
    });

    const images = imageDataMap[claimId] || [];
    if (images.length > 0) {
      excelRow.height = 60;
      
      const maxImages = 5;
      for (let j = 0; j < Math.min(images.length, maxImages); j++) {
        const imageData = images[j];
        if (imageData && imageData.base64) {
          try {
            const imageId = workbook.addImage({ base64: imageData.base64, extension: 'jpeg' });
            const columnIdx = 11 + j; // Starts at column L (index 11)

            worksheet.addImage(imageId, {
              tl: { col: columnIdx, row: excelRow.number - 1 },
              br: { col: columnIdx + 1, row: excelRow.number }
            });
          } catch (imageError) {
            console.error(`Failed to add image ${imageData.name} to Excel:`, imageError);
          }
        }
      }
    }
  }

  return workbook;
}
  /**
   * Main function to scrape the table, process images, and generate the Excel file.
   */
  async function exportToExcel() {
    console.log("=== Starting Export Process ===");
    
    const table = document.querySelector("form#listForm table");
    if (!table) {
      alert("Error: Could not find the data table on the page.");
      return;
    }

    const rows = table.querySelectorAll("tr");
    const excelData = [];
    const imageDataMap = {};

    // Add header row to our data structure
    excelData.push(["Claim ID", "Warehouse", "Tracking #", "QTY", "Reference #", "Valid Time", "Claim Time", "Created", "Completed", "Updated", "Notes", "Images"]);

    const progressDiv = document.createElement('div');
    progressDiv.id = 'export-progress-indicator';
    progressDiv.style.cssText = `position: fixed; top: 70px; right: 20px; z-index: 10001; background-color: #2196F3; color: white; padding: 15px 25px; font-size: 16px; border-radius: 5px; box-shadow: 0 4px 8px rgba(0,0,0,0.3);`;
    document.body.appendChild(progressDiv);

    const totalRows = rows.length - 1; // Exclude header row

    // Iterate over data rows (skip the table header row at index 0)
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const checkbox = row.querySelector('input.checkItem');

      // If a checkbox doesn't exist or is NOT checked, skip to the next row.
      if (!checkbox || !checkbox.checked) {
        continue;
      }
      const cells = row.querySelectorAll("td");
      if (cells.length < 8) continue; // Skip if row doesn't have enough cells

      const mainCell = cells[2];
      const claimId = mainCell.querySelector("b")?.innerText.trim();
      if (!claimId) continue;

      progressDiv.textContent = `Processing row ${i}/${totalRows} (ID: ${claimId})...`;
      
      const warehouse = mainCell.querySelectorAll("b")[1]?.innerText.trim() || "";
      const tracking = mainCell.querySelector(".tracking")?.innerText.trim() || "";
      const qty = mainCell.innerHTML.match(/包裹数量：<\/span><b>(\d+)<\/b>/)?.[1] || "";
      const ref = mainCell.innerHTML.match(/参考号：<\/span><b>(.*?)<\/b>/)?.[1] || "";
      const validTime = cells[5]?.innerText.match(/有效时间：(.+)/)?.[1].trim() || "";
      const claimTime = cells[5]?.innerText.match(/认领时间：(.+)/)?.[1].trim() || "";
      const created = cells[6]?.innerText.match(/创建时间：(.+)/)?.[1].trim() || "";
      const completed = cells[6]?.innerText.match(/完成时间：(.+)/)?.[1].trim() || "";
      const updated = cells[6]?.innerText.match(/更新时间：(.+)/)?.[1].trim() || "";
      const notes = cells[7]?.innerText.trim().replace(/\n/g, " / ") || "";

      // Process images for this row
      const imageData = await processRowImages(row, claimId);
      if (imageData.length > 0) {
        imageDataMap[claimId] = imageData;
      }

      excelData.push([claimId, warehouse, tracking, qty, ref, validTime, claimTime, created, completed, updated, notes, ""]);
      
      // Small delay to avoid overwhelming the server with popup requests
      await new Promise(resolve => setTimeout(resolve, 500));
    }

    progressDiv.textContent = 'Generating Excel file... This may take a moment.';
    progressDiv.style.backgroundColor = '#FFC107';

    try {
      const workbook = await createExcelWorkbook(excelData, imageDataMap);
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.setAttribute("href", url);
      link.setAttribute("download", `Package_Data_${new Date().toISOString().slice(0, 10)}.xlsx`);
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
      
      progressDiv.textContent = `Success! Exported ${totalRows} rows.`;
      progressDiv.style.backgroundColor = '#4CAF50';
    } catch (error) {
      console.error("Fatal error during Excel file generation:", error);
      alert("An error occurred while creating the Excel file. Check the console for details.");
      progressDiv.textContent = 'Export Failed!';
      progressDiv.style.backgroundColor = '#f44336';
    }
    
    setTimeout(() => {
      progressDiv.remove();
    }, 5000);
  }

  /**
   * Creates and injects the "Export to Excel" button onto the page.
   */
  function createExportButton() {
    if (document.getElementById('wms-export-excel-btn')) return; // Avoid creating duplicate buttons

    const button = document.createElement('button');
    button.id = 'wms-export-excel-btn';
    button.innerText = 'Export to Excel';
    button.style.cssText = `position: fixed; top: 20px; right: 20px; z-index: 10000; background-color: #4CAF50; color: white; border: none; padding: 10px 20px; font-size: 14px; border-radius: 5px; cursor: pointer; box-shadow: 0 2px 5px rgba(0,0,0,0.2);`;
    button.onclick = exportToExcel;
    document.body.appendChild(button);
  }

  /**
   * Initializes the script once the page is ready.
   */
  function initialize() {
    // We are on the correct page if the main form exists.
    if (document.querySelector("form#listForm")) {
      console.log("WMS Exporter: Target form found. Initializing button.");
      createExportButton();
    } else {
      console.log("WMS Exporter: Target form not found. Script will not activate.");
    }
  }

  // Run the initialization logic after the DOM is fully loaded.
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initialize);
  } else {
    initialize();
  }
  
})();