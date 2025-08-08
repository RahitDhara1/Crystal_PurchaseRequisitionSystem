// ...existing code...
function getHeaderMap(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const headerMap = {};
  headers.forEach((header, i) => {
    headerMap[header] = i;
  });
  return headerMap;
}

/**
 * Fetches details for a given Vendor ID from the Vendor_Master sheet.
 * @param {string} vendorId The Vendor ID.
 * @returns {object} An object containing vendor details.
 */
function getVendorDetails(vendorId) {
  if (!vendorId) return {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const vendorSheet = ss.getSheetByName('Vendor_Master');
  if (!vendorSheet) {
    Logger.log("Sheet 'Vendor_Master' not found.");
    return {};
  }

  const vendorHeaders = getHeaderMap(vendorSheet);
  const vendorData = vendorSheet.getRange(2, 1, vendorSheet.getLastRow() - 1, vendorSheet.getLastColumn()).getValues();

  for (const vendorRow of vendorData) {
    if (vendorRow[vendorHeaders['Vendor ID']] === vendorId) {
      return {
        id: vendorRow[vendorHeaders['Vendor ID']],
        companyName: vendorRow[vendorHeaders['Company Name']],
        contactPerson: vendorRow[vendorHeaders['Contact Person Name']],
        contactNumber: vendorRow[vendorHeaders['Contact Number']],
        email: vendorRow[vendorHeaders['Email ID']],
        // Add other vendor fields as needed
      };
    }
  }
  return {};
}

// Function to get data for the requisition approval page
function getPendingRequisitions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reqSheet = ss.getSheetByName('Requisitions');
  const itemsSheet = ss.getSheetByName('Items');

  if (!reqSheet || !itemsSheet) {
    throw new Error("Required sheets ('Requisitions', 'Items') not found.");
  }

  const reqHeaders = getHeaderMap(reqSheet);
  const itemHeaders = getHeaderMap(itemsSheet);
  const reqData = reqSheet.getRange(2, 1, reqSheet.getLastRow() - 1, reqSheet.getLastColumn()).getValues();
  const itemsData = itemsSheet.getRange(2, 1, itemsSheet.getLastRow() - 1, itemsSheet.getLastColumn()).getValues();

  const pendingRequisitions = reqData
    .filter(row => row[reqHeaders['Current Status']] === 'Submitted')
    .map(reqRow => {
      const prID = reqRow[reqHeaders['Requisition ID']];
      const vendorId = reqRow[reqHeaders['Vendor ID']];
      const vendorDetails = getVendorDetails(vendorId);

      const itemsForReq = itemsData
        .filter(itemRow => itemRow[itemHeaders['Requisition ID']] === prID)
        .map(itemRow => ({
          itemName: itemRow[itemHeaders['Item Name']],
          quantity: itemRow[itemHeaders['Quantity Required']],
          uom: itemRow[itemHeaders['UOM']],
          rate: itemRow[itemHeaders['Rate']],
          totalValue: itemRow[itemHeaders['Total Value (Incl. GST)']]
        }));

      return {
        id: prID,
        date: reqRow[reqHeaders['Date of Requisition']],
        requestedBy: reqRow[reqHeaders['Requested By']],
        site: reqRow[reqHeaders['Site']],
        status: reqRow[reqHeaders['Current Status']],
        expectedDeliveryDate: reqRow[reqHeaders['Expected Delivery Date']],
        totalValue: reqRow[reqHeaders['Total Value Incl. GST']],
        vendor: vendorDetails,
        items: itemsForReq
      };
    });

  return { pendingRequisitions: JSON.stringify(pendingRequisitions) };
}


// Function to get data for the PO approval page
// ...existing code...
// Function to process approval/rejection from the approval page
function processApproval(prID, action, remarks) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reqSheet = ss.getSheetByName('Requisitions');
  const headers = getHeaderMap(reqSheet);
  const data = reqSheet.getRange(2, 1, reqSheet.getLastRow(), reqSheet.getLastColumn()).getValues();

  const userEmail = Session.getActiveUser().getEmail();
  const timestamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss");

  for (let i = 0; i < data.length; i++) {
    if (data[i][headers['Requisition ID']] === prID) {
      const rowNum = i + 2;
      const newStatus = action === 'Approved' ? 'Requisition Approved' : 'Requisition Rejected';
      
      reqSheet.getRange(rowNum, headers['Current Status'] + 1).setValue(newStatus);
      reqSheet.getRange(rowNum, headers['Approval Status'] + 1).setValue(action);
      reqSheet.getRange(rowNum, headers['Approver Remarks'] + 1).setValue(remarks);
      
      // Optionally, log who approved/rejected and when
      // This assumes you have columns for this, which are not in the provided list.
      // If you add them, you can uncomment and adapt the following lines.
      // reqSheet.getRange(rowNum, headers['Approved By'] + 1).setValue(userEmail);
      // reqSheet.getRange(rowNum, headers['Approval Timestamp'] + 1).setValue(timestamp);

      return { success: true, prID: prID, status: newStatus };
    }
  }
  throw new Error(`Requisition with ID ${prID} not found.`);
}

// ...existing code...
function getRequisitionDetailsForPO(prID) {
  if (!prID) {
    throw new Error("Requisition ID is required.");
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reqSheet = ss.getSheetByName('Requisitions');
  const itemsSheet = ss.getSheetByName('Items');

  if (!reqSheet || !itemsSheet) {
    throw new Error("Required sheets ('Requisitions', 'Items') not found.");
  }

  const reqHeaders = getHeaderMap(reqSheet);
  const itemHeaders = getHeaderMap(itemsSheet);
  const reqData = reqSheet.getRange(2, 1, reqSheet.getLastRow() - 1, reqSheet.getLastColumn()).getValues();
  const itemsData = itemsSheet.getRange(2, 1, itemsSheet.getLastRow() - 1, itemsSheet.getLastColumn()).getValues();

  let requisitionDetails = {};
  let found = false;

  for (const row of reqData) {
    if (row[reqHeaders['Requisition ID']] === prID) {
      const expectedDate = row[reqHeaders['Expected Delivery Date']];
      const vendorId = row[reqHeaders['Vendor ID']];
      const vendorDetails = getVendorDetails(vendorId);
      
      requisitionDetails = {
        paymentTerms: row[reqHeaders['Payment Terms']],
        deliveryTerms: row[reqHeaders['Delivery Terms']],
        expectedDeliveryDate: expectedDate ? new Date(expectedDate).toISOString().split('T')[0] : '',
        site: row[reqHeaders['Site']] || '',
        vendor: vendorDetails // Pass the whole vendor object
      };
      found = true;
      break;
    }
  }

  if (!found) {
    throw new Error(`Requisition ${prID} not found.`);
  }

  const itemsForReq = itemsData
    .filter(itemRow => itemRow[itemHeaders['Requisition ID']] === prID)
    .map(itemRow => ({
      itemName: itemRow[itemHeaders['Item Name']],
      purpose: itemRow[itemHeaders['Purpose / Application']],
      quantity: itemRow[itemHeaders['Quantity Required']],
      uom: itemRow[itemHeaders['UOM']],
      rate: itemRow[itemHeaders['Rate']],
      gst: itemRow[itemHeaders['GST']],
      totalCost: itemRow[itemHeaders['Total Value (Incl. GST)']] 
    }));
  
  requisitionDetails.items = itemsForReq;
  return requisitionDetails;
}

// Handle PO form submission
// ...existing code...
// Handle form submission
function submitPR(formData, lineItems) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reqSheet = ss.getSheetByName('Requisitions');
  const itemsSheet = ss.getSheetByName('Items');

  const site = formData.site.replace(/\s+/g, '');
  const { serial, monthYear } = getNextPRSerial(site);
  const prID = `PR-${site}-${monthYear}/${serial}`;

  // Calculate total value from line items
  const totalValue = lineItems.reduce((sum, item) => sum + parseFloat(item.totalValue || 0), 0);

  const now = new Date();
  const timestamp = Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss");
  const requisitionDate = Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");

  // Save master details in the correct column order
  const requisitionData = [
    timestamp, // Timestamp
    requisitionDate, // Date of Requisition
    prID, // Requisition ID
    'Submitted', // Approval Status
    'Submitted', // Current Status
    '', // Approver Remarks
    totalValue.toFixed(2), // Total Value Incl. GST
    formData.purchaseCategory, // Purchase Category
    formData.paymentTerms, // Payment Terms
    formData.deliveryTerms, // Delivery Terms
    Session.getActiveUser().getEmail(), // Requested By
    formData.site, // Site
    formData.deliveryLocation, // Delivery Location
    formData.isVendorRegistered, // Is the Vendor Registered with Us?
    formData.vendorId, // Vendor ID
    formData.quotationPI, // Upload Quotation / PI Final Agreed
    formData.supportingDocs, // Upload Supporting Documents
    formData.isCustomerReimbursable, // Is the customer reimbursable?
    '', // Approved PR Link
    '', // PDF Link
    '', // Approval Link(View)
    Session.getActiveUser().getEmail(), // Email Address
    formData.expectedDeliveryDate // Expected Delivery Date
  ];
  
  reqSheet.appendRow(requisitionData);

  // Save line items
  lineItems.forEach(item => {
    itemsSheet.appendRow([
      prID,
      item.itemName,
      item.purpose,
      item.quantity,
      item.uom,
      item.rate,
      item.gst,
      item.totalValue
    ]);
  });

  return { success: true, prID: prID };
}
// ...existing code...
```

### 2. `requisitionApproval.js` - Approval Page UI

This file needs to be updated to correctly display the vendor information which is now nested inside a `vendor` object.

````javascript
// filepath: c:\Users\Rahit\Desktop\CrystalGroup\requisitionApproval.js
// ...existing code...
        function showModal(prID) {
            const req = requisitions.find(r => r.id === prID);
            if (!req) return;

            document.getElementById('modal-title').innerText = 'Requisition Details: ' + prID;
            
// ...existing code...
            // Vendor Info
            document.getElementById('modal-vendor-info').innerHTML = `
                <h4>Vendor Information</h4>
                <p><strong>Company Name:</strong> ${req.vendor.companyName || 'N/A'}</p>
                <p><strong>Contact Person:</strong> ${req.vendor.contactPerson || 'N/A'}</p>
                <p><strong>Email:</strong> ${req.vendor.email || 'N/A'}</p>
                <p><strong>Phone:</strong> ${req.vendor.contactNumber || 'N/A'}</p>
            `;

            // Items Table
            const itemsTbody = document.getElementById('modal-items-table');
// ...existing code...
```

### 3. `allRequisitionPage.js` - All Requisitions View

This page displays a summary. Since the `Vendor ID` is now in the main sheet instead of the name, I'll update the `getAllRequisitions` function in `Code.js` to fetch and include the vendor name for display.

First, add the `getAllRequisitions` function to `Code.js`:

````javascript
// filepath: c:\Users\Rahit\Desktop\CrystalGroup\Code.js
// ...existing code...
// Add this function to your Code.js file

function getAllRequisitions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reqSheet = ss.getSheetByName('Requisitions');
  if (!reqSheet) {
    Logger.log("Sheet 'Requisitions' not found.");
    return [];
  }

  const reqHeaders = getHeaderMap(reqSheet);
  const reqData = reqSheet.getRange(2, 1, reqSheet.getLastRow() - 1, reqSheet.getLastColumn()).getValues();

  const allRequisitions = reqData.map(row => {
    const vendorId = row[reqHeaders['Vendor ID']];
    const vendorDetails = getVendorDetails(vendorId); // Reuse the helper function

    // Create a serializable object for the frontend
    return {
      'Requisition ID': row[reqHeaders['Requisition ID']],
      'Date of Requisition': row[reqHeaders['Date of Requisition']] ? new Date(row[reqHeaders['Date of Requisition']]).toLocaleDateString() : 'N/A',
      'Requested By': row[reqHeaders['Requested By']],
      'Site': row[reqHeaders['Site']],
      'Current Status': row[reqHeaders['Current Status']],
      'Total Value': row[reqHeaders['Total Value Incl. GST']],
      'Vendor Name': vendorDetails.companyName || 'N/A', // Add vendor name
      'Vendor ID': vendorId,
      'Payment Terms': row[reqHeaders['Payment Terms']],
      'Delivery Terms': row[reqHeaders['Delivery Terms']],
      'Delivery Location': row[reqHeaders['Delivery Location']]
    };
  });

  return allRequisitions;
}

// ...existing code...
```

Now, update `allRequisitionPage.js` to display the vendor name.

````javascript
// filepath: c:\Users\Rahit\Desktop\CrystalGroup\allRequisitionPage.js
// ...existing code...
      container.innerHTML = data.map((req, idx) => `
        <div class="requisition-card" onclick="showReqDetails(${idx})">
          <strong>PR ID:</strong> ${req['Requisition ID'] || ''}<br>
          <strong>Requested By:</strong> ${req['Requested By'] || ''}<br>
          <strong>Vendor:</strong> ${req['Vendor Name'] || ''}<br>
          <strong>Site:</strong> ${req['Site'] || ''}<br>
          <strong>Status:</strong> ${req['Current Status'] || ''}<br>
          <strong>Date:</strong> ${req['Date of Requisition'] || ''}
        </div>
      `).join('');
      window.allReqData = data;
    }).getAllRequisitions();

    window.showReqDetails = function(idx) {
// ...existing code...
```

These changes should align your application with your new spreadsheet structure. You will also need to create and populate the `Vendor_Master` sheet with the necessary vendor details, including a unique `Vendor ID` for each vendor. The `requisitionForm.js` may also need changes to allow users to select a vendor (and get their ID) rather than manually entering vendor details.// filepath: c:\Users\Rahit\Desktop\CrystalGroup\Code.js
// ...existing code...
function getHeaderMap(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const headerMap = {};
  headers.forEach((header, i) => {
    headerMap[header] = i;
  });
  return headerMap;
}

/**
 * Fetches details for a given Vendor ID from the Vendor_Master sheet.
 * @param {string} vendorId The Vendor ID.
 * @returns {object} An object containing vendor details.
 */
function getVendorDetails(vendorId) {
  if (!vendorId) return {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const vendorSheet = ss.getSheetByName('Vendor_Master');
  if (!vendorSheet) {
    Logger.log("Sheet 'Vendor_Master' not found.");
    return {};
  }

  const vendorHeaders = getHeaderMap(vendorSheet);
  const vendorData = vendorSheet.getRange(2, 1, vendorSheet.getLastRow() - 1, vendorSheet.getLastColumn()).getValues();

  for (const vendorRow of vendorData) {
    if (vendorRow[vendorHeaders['Vendor ID']] === vendorId) {
      return {
        id: vendorRow[vendorHeaders['Vendor ID']],
        companyName: vendorRow[vendorHeaders['Company Name']],
        contactPerson: vendorRow[vendorHeaders['Contact Person Name']],
        contactNumber: vendorRow[vendorHeaders['Contact Number']],
        email: vendorRow[vendorHeaders['Email ID']],
        // Add other vendor fields as needed
      };
    }
  }
  return {};
}

// Function to get data for the requisition approval page
function getPendingRequisitions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reqSheet = ss.getSheetByName('Requisitions');
  const itemsSheet = ss.getSheetByName('Items');

  if (!reqSheet || !itemsSheet) {
    throw new Error("Required sheets ('Requisitions', 'Items') not found.");
  }

  const reqHeaders = getHeaderMap(reqSheet);
  const itemHeaders = getHeaderMap(itemsSheet);
  const reqData = reqSheet.getRange(2, 1, reqSheet.getLastRow() - 1, reqSheet.getLastColumn()).getValues();
  const itemsData = itemsSheet.getRange(2, 1, itemsSheet.getLastRow() - 1, itemsSheet.getLastColumn()).getValues();

  const pendingRequisitions = reqData
    .filter(row => row[reqHeaders['Current Status']] === 'Submitted')
    .map(reqRow => {
      const prID = reqRow[reqHeaders['Requisition ID']];
      const vendorId = reqRow[reqHeaders['Vendor ID']];
      const vendorDetails = getVendorDetails(vendorId);

      const itemsForReq = itemsData
        .filter(itemRow => itemRow[itemHeaders['Requisition ID']] === prID)
        .map(itemRow => ({
          itemName: itemRow[itemHeaders['Item Name']],
          quantity: itemRow[itemHeaders['Quantity Required']],
          uom: itemRow[itemHeaders['UOM']],
          rate: itemRow[itemHeaders['Rate']],
          totalValue: itemRow[itemHeaders['Total Value (Incl. GST)']]
        }));

      return {
        id: prID,
        date: reqRow[reqHeaders['Date of Requisition']],
        requestedBy: reqRow[reqHeaders['Requested By']],
        site: reqRow[reqHeaders['Site']],
        status: reqRow[reqHeaders['Current Status']],
        expectedDeliveryDate: reqRow[reqHeaders['Expected Delivery Date']],
        totalValue: reqRow[reqHeaders['Total Value Incl. GST']],
        vendor: vendorDetails,
        items: itemsForReq
      };
    });

  return { pendingRequisitions: JSON.stringify(pendingRequisitions) };
}


// Function to get data for the PO approval page
// ...existing code...
// Function to process approval/rejection from the approval page
function processApproval(prID, action, remarks) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reqSheet = ss.getSheetByName('Requisitions');
  const headers = getHeaderMap(reqSheet);
  const data = reqSheet.getRange(2, 1, reqSheet.getLastRow(), reqSheet.getLastColumn()).getValues();

  const userEmail = Session.getActiveUser().getEmail();
  const timestamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss");

  for (let i = 0; i < data.length; i++) {
    if (data[i][headers['Requisition ID']] === prID) {
      const rowNum = i + 2;
      const newStatus = action === 'Approved' ? 'Requisition Approved' : 'Requisition Rejected';
      
      reqSheet.getRange(rowNum, headers['Current Status'] + 1).setValue(newStatus);
      reqSheet.getRange(rowNum, headers['Approval Status'] + 1).setValue(action);
      reqSheet.getRange(rowNum, headers['Approver Remarks'] + 1).setValue(remarks);
      
      // Optionally, log who approved/rejected and when
      // This assumes you have columns for this, which are not in the provided list.
      // If you add them, you can uncomment and adapt the following lines.
      // reqSheet.getRange(rowNum, headers['Approved By'] + 1).setValue(userEmail);
      // reqSheet.getRange(rowNum, headers['Approval Timestamp'] + 1).setValue(timestamp);

      return { success: true, prID: prID, status: newStatus };
    }
  }
  throw new Error(`Requisition with ID ${prID} not found.`);
}

// ...existing code...
function getRequisitionDetailsForPO(prID) {
  if (!prID) {
    throw new Error("Requisition ID is required.");
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reqSheet = ss.getSheetByName('Requisitions');
  const itemsSheet = ss.getSheetByName('Items');

  if (!reqSheet || !itemsSheet) {
    throw new Error("Required sheets ('Requisitions', 'Items') not found.");
  }

  const reqHeaders = getHeaderMap(reqSheet);
  const itemHeaders = getHeaderMap(itemsSheet);
  const reqData = reqSheet.getRange(2, 1, reqSheet.getLastRow() - 1, reqSheet.getLastColumn()).getValues();
  const itemsData = itemsSheet.getRange(2, 1, itemsSheet.getLastRow() - 1, itemsSheet.getLastColumn()).getValues();

  let requisitionDetails = {};
  let found = false;

  for (const row of reqData) {
    if (row[reqHeaders['Requisition ID']] === prID) {
      const expectedDate = row[reqHeaders['Expected Delivery Date']];
      const vendorId = row[reqHeaders['Vendor ID']];
      const vendorDetails = getVendorDetails(vendorId);
      
      requisitionDetails = {
        paymentTerms: row[reqHeaders['Payment Terms']],
        deliveryTerms: row[reqHeaders['Delivery Terms']],
        expectedDeliveryDate: expectedDate ? new Date(expectedDate).toISOString().split('T')[0] : '',
        site: row[reqHeaders['Site']] || '',
        vendor: vendorDetails // Pass the whole vendor object
      };
      found = true;
      break;
    }
  }

  if (!found) {
    throw new Error(`Requisition ${prID} not found.`);
  }

  const itemsForReq = itemsData
    .filter(itemRow => itemRow[itemHeaders['Requisition ID']] === prID)
    .map(itemRow => ({
      itemName: itemRow[itemHeaders['Item Name']],
      purpose: itemRow[itemHeaders['Purpose / Application']],
      quantity: itemRow[itemHeaders['Quantity Required']],
      uom: itemRow[itemHeaders['UOM']],
      rate: itemRow[itemHeaders['Rate']],
      gst: itemRow[itemHeaders['GST']],
      totalCost: itemRow[itemHeaders['Total Value (Incl. GST)']] 
    }));
  
  requisitionDetails.items = itemsForReq;
  return requisitionDetails;
}

// Handle PO form submission
// ...existing code...
// Handle form submission
function submitPR(formData, lineItems) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reqSheet = ss.getSheetByName('Requisitions');
  const itemsSheet = ss.getSheetByName('Items');

  const site = formData.site.replace(/\s+/g, '');
  const { serial, monthYear } = getNextPRSerial(site);
  const prID = `PR-${site}-${monthYear}/${serial}`;

  // Calculate total value from line items
  const totalValue = lineItems.reduce((sum, item) => sum + parseFloat(item.totalValue || 0), 0);

  const now = new Date();
  const timestamp = Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss");
  const requisitionDate = Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");

  // Save master details in the correct column order
  const requisitionData = [
    timestamp, // Timestamp
    requisitionDate, // Date of Requisition
    prID, // Requisition ID
    'Submitted', // Approval Status
    'Submitted', // Current Status
    '', // Approver Remarks
    totalValue.toFixed(2), // Total Value Incl. GST
    formData.purchaseCategory, // Purchase Category
    formData.paymentTerms, // Payment Terms
    formData.deliveryTerms, // Delivery Terms
    Session.getActiveUser().getEmail(), // Requested By
    formData.site, // Site
    formData.deliveryLocation, // Delivery Location
    formData.isVendorRegistered, // Is the Vendor Registered with Us?
    formData.vendorId, // Vendor ID
    formData.quotationPI, // Upload Quotation / PI Final Agreed
    formData.supportingDocs, // Upload Supporting Documents
    formData.isCustomerReimbursable, // Is the customer reimbursable?
    '', // Approved PR Link
    '', // PDF Link
    '', // Approval Link(View)
    Session.getActiveUser().getEmail(), // Email Address
    formData.expectedDeliveryDate // Expected Delivery Date
  ];
  
  reqSheet.appendRow(requisitionData);

  // Save line items
  lineItems.forEach(item => {
    itemsSheet.appendRow([
      prID,
      item.itemName,
      item.purpose,
      item.quantity,
      item.uom,
      item.rate,
      item.gst,
      item.totalValue
    ]);
  });

  return { success: true, prID: prID };
}
// ...existing code...
```

### 2. `requisitionApproval.js` - Approval Page UI

This file needs to be updated to correctly display the vendor information which is now nested inside a `vendor` object.

````javascript
// filepath: c:\Users\Rahit\Desktop\CrystalGroup\requisitionApproval.js
// ...existing code...
        function showModal(prID) {
            const req = requisitions.find(r => r.id === prID);
            if (!req) return;

            document.getElementById('modal-title').innerText = 'Requisition Details: ' + prID;
            
// ...existing code...
            // Vendor Info
            document.getElementById('modal-vendor-info').innerHTML = `
                <h4>Vendor Information</h4>
                <p><strong>Company Name:</strong> ${req.vendor.companyName || 'N/A'}</p>
                <p><strong>Contact Person:</strong> ${req.vendor.contactPerson || 'N/A'}</p>
                <p><strong>Email:</strong> ${req.vendor.email || 'N/A'}</p>
                <p><strong>Phone:</strong> ${req.vendor.contactNumber || 'N/A'}</p>
            `;

            // Items Table
            const itemsTbody = document.getElementById('modal-items-table');
// ...existing code...
```

### 3. `allRequisitionPage.js` - All Requisitions View

This page displays a summary. Since the `Vendor ID` is now in the main sheet instead of the name, I'll update the `getAllRequisitions` function in `Code.js` to fetch and include the vendor name for display.

First, add the `getAllRequisitions` function to `Code.js`:

````javascript
// filepath: c:\Users\Rahit\Desktop\CrystalGroup\Code.js
// ...existing code...
// Add this function to your Code.js file

function getAllRequisitions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reqSheet = ss.getSheetByName('Requisitions');
  if (!reqSheet) {
    Logger.log("Sheet 'Requisitions' not found.");
    return [];
  }

  const reqHeaders = getHeaderMap(reqSheet);
  const reqData = reqSheet.getRange(2, 1, reqSheet.getLastRow() - 1, reqSheet.getLastColumn()).getValues();

  const allRequisitions = reqData.map(row => {
    const vendorId = row[reqHeaders['Vendor ID']];
    const vendorDetails = getVendorDetails(vendorId); // Reuse the helper function

    // Create a serializable object for the frontend
    return {
      'Requisition ID': row[reqHeaders['Requisition ID']],
      'Date of Requisition': row[reqHeaders['Date of Requisition']] ? new Date(row[reqHeaders['Date of Requisition']]).toLocaleDateString() : 'N/A',
      'Requested By': row[reqHeaders['Requested By']],
      'Site': row[reqHeaders['Site']],
      'Current Status': row[reqHeaders['Current Status']],
      'Total Value': row[reqHeaders['Total Value Incl. GST']],
      'Vendor Name': vendorDetails.companyName || 'N/A', // Add vendor name
      'Vendor ID': vendorId,
      'Payment Terms': row[reqHeaders['Payment Terms']],
      'Delivery Terms': row[reqHeaders['Delivery Terms']],
      'Delivery Location': row[reqHeaders['Delivery Location']]
    };
  });

  return allRequisitions;
}

// ...existing code...
```

Now, update `allRequisitionPage.js` to display the vendor name.

````javascript
// filepath: c:\Users\Rahit\Desktop\CrystalGroup\allRequisitionPage.js
// ...existing code...
      container.innerHTML = data.map((req, idx) => `
        <div class="requisition-card" onclick="showReqDetails(${idx})">
          <strong>PR ID:</strong> ${req['Requisition ID'] || ''}<br>
          <strong>Requested By:</strong> ${req['Requested By'] || ''}<br>
          <strong>Vendor:</strong> ${req['Vendor Name'] || ''}<br>
          <strong>Site:</strong> ${req['Site'] || ''}<br>
          <strong>Status:</strong> ${req['Current Status'] || ''}<br>
          <strong>Date:</strong> ${req['Date of Requisition'] || ''}
        </div>
      `).join('');
      window.allReqData = data;
    }).getAllRequisitions();

    window.showReqDetails = function(idx) {
// ...existing code...
```

These changes should align your application with your new spreadsheet structure. You will also need to create and populate the `Vendor_Master` sheet with the necessary vendor details, including a unique `Vendor ID` for each vendor. The `requisitionForm.js` may also need changes to allow users to select a vendor (and get their ID) rather than manually entering vendor details.