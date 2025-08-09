function doGet(e) {
  const page = e.parameter.page || 'home';
  const user = {
    email: Session.getActiveUser().getEmail(),
    name: Session.getActiveUser().getEmail().split('@')[0]
  };

  let template;
  let pageContent = '';
  let pageTitle = 'Purchase Requisition System';

  if (page === 'home') {
    pageTitle = 'Dashboard';
    template = HtmlService.createTemplateFromFile('home');
    template.stats = getDashboardStats();
    pageContent = template.evaluate().getContent();
  } else if (page === 'requisitionForm') {
    pageTitle = 'New Requisition';
    pageContent = requisitionFormPage(user);
  } else if (page === 'approval') {
    pageTitle = 'Requisition Approval';
    const data = getPendingRequisitions();
    pageContent = requisitionApprovalPage(data);
  } else if (page === 'poForm') {
    pageTitle = 'Create Purchase Order';
    pageContent = poFormPage(user);
  } else if (page === 'poApproval') {
    pageTitle = 'PO Approval';
    const data = getPendingPOs();
    pageContent = poApprovalPage(data);
  } else if (page === 'allRequisitions') {
    pageTitle = 'All Requisitions';
    pageContent = allRequisitionsPage();
  } else {
    template = HtmlService.createTemplateFromFile('home');
    template.stats = getDashboardStats();
    pageContent = template.evaluate().getContent();
  }

  const layoutTemplate = HtmlService.createTemplateFromFile('layout');
  layoutTemplate.activePage = page;
  layoutTemplate.pageTitle = pageTitle;
  layoutTemplate.user = user;
  layoutTemplate.pageContent = pageContent;
  
  return layoutTemplate.evaluate().setTitle(pageTitle);
}

/**
 * Creates a map of column headers to their zero-based index.
 * @param {Sheet} sheet The Google Sheet to get headers from.
 * @returns {Object} An object where keys are header names and values are column indices.
 */
// function getHeaderMap(sheet) {
//   const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
//   const headerMap = {};
//   headers.forEach((header, i) => {
//     headerMap[header] = i;
//   });
//   return headerMap;
// }
function getHeaderMap(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const headerMap = {};
  headers.forEach((header, i) => {
    headerMap[header] = i;
  });
  return headerMap;
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Updates the status of a given requisition ID in the 'Requisitions' sheet.
 * @param {string} prID The Requisition ID.
 * @param {string} status The new status to set.
 */
function updateRequisitionStatus(prID, status) {
  if (!prID || !status) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reqSheet = ss.getSheetByName('Requisitions');
  const headers = getHeaderMap(reqSheet);
  const dataRange = reqSheet.getRange(2, 1, reqSheet.getLastRow() - 1, reqSheet.getLastColumn());
  const data = dataRange.getValues();

  for (let i = 0; i < data.length; i++) {
    if (data[i][headers['Requisition ID']] === prID) {
      reqSheet.getRange(i + 2, headers['Current Status'] + 1).setValue(status);
      break;
    }
  }
}



// Function to get data for the approval page
// function getPendingRequisitions() {
//   const cache = CacheService.getScriptCache();
//   // Clear cache to ensure fresh data is fetched
//   cache.remove('pendingRequisitions');
//   const cached = cache.get('pendingRequisitions');
//   if (cached) {
//     return JSON.parse(cached);
//   }

//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const reqSheet = ss.getSheetByName('Requisitions');
//   const itemsSheet = ss.getSheetByName('Items');
  
//   if (!reqSheet || !itemsSheet) {
//     return { pendingRequisitions: '[]' };
//   }

//   const reqHeaders = getHeaderMap(reqSheet);
//   const itemHeaders = getHeaderMap(itemsSheet);
//   const reqData = reqSheet.getRange(2, 1, reqSheet.getLastRow() - 1, reqSheet.getLastColumn()).getValues();
//   const itemsData = itemsSheet.getRange(2, 1, itemsSheet.getLastRow() - 1, itemsSheet.getLastColumn()).getValues();

//   // Only show those with Current Status = 'Requisition Submitted'
//   const pendingRequisitions = reqData
//     .map((row, index) => ({ data: row, rowIndex: index + 2 }))
//     .filter(r => r.data[reqHeaders['Current Status']] === 'Requisition Submitted')
//     .map(r => {
//       const reqRow = r.data;
//       const prID = reqRow[reqHeaders['Requisition ID']];
      
//       const itemsForReq = itemsData
//         .filter(itemRow => itemRow[itemHeaders['Requisition ID']] === prID)
//         .map(itemRow => ({
//           itemName: itemRow[itemHeaders['Item Name']],
//           purpose: itemRow[itemHeaders['Purpose / Application']],
//           quantity: itemRow[itemHeaders['Quantity Required']],
//           uom: itemRow[itemHeaders['Rate']],
//           totalValue: itemRow[itemHeaders['Total Value (Incl. GST)']]
//         }));

//       const vendor = {
//         name: reqRow[reqHeaders['Registered vendor company name']] || reqRow[reqHeaders['vendor company name']],
//         contactPerson: reqRow[reqHeaders['Registered Vendor Contact Person Name']] || reqRow[reqHeaders['Vendor Contact Person Name']],
//         contactNumber: reqRow[reqHeaders['Registered Vendor Contact Person Number']] || reqRow[reqHeaders['Vendor Contact Person Number']],
//         email: reqRow[reqHeaders['Registered Vendor Email ID']] || reqRow[reqHeaders['Vendor Email ID']]
//       };

//       Logger.log('Vendor for PR %s: %s', prID, vendor.name);

//       return {
//         id: prID,
//         date: reqRow[reqHeaders['Date of Requisition']],
//         expectedDeliveryDate: reqRow[reqHeaders['Expected Delivery Date']],
//         status: reqRow[reqHeaders['Approval Status']],
//         totalValue: reqRow[reqHeaders['Total Value Incl. GST']],
//         purchaseCategory: reqRow[reqHeaders['Purchase Category']],
//         requestedBy: reqRow[reqHeaders['Requested By']],
//         site: reqRow[reqHeaders['Site']],
//         vendor: vendor,
//         items: itemsForReq,
//         rowIndex: r.rowIndex
//       };
//     });

//   const result = { pendingRequisitions: JSON.stringify(pendingRequisitions) };
//   cache.put('pendingRequisitions', JSON.stringify(result), 120); 
//   return result;
// }
// function getPendingRequisitions() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const reqSheet = ss.getSheetByName('Requisitions');
//   const itemsSheet = ss.getSheetByName('Items');

//   if (!reqSheet || !itemsSheet) {
//     throw new Error("Required sheets ('Requisitions', 'Items') not found.");
//   }

//   const reqHeaders = getHeaderMap(reqSheet);
//   const itemHeaders = getHeaderMap(itemsSheet);

//   const reqData = reqSheet.getRange(2, 1, reqSheet.getLastRow() - 1, reqSheet.getLastColumn()).getValues();
//   const itemsData = itemsSheet.getRange(2, 1, itemsSheet.getLastRow() - 1, itemsSheet.getLastColumn()).getValues();

//   const pendingRequisitions = reqData
//     .map((row, index) => ({ data: row, rowIndex: index + 2 }))
//     .filter(r => r.data[reqHeaders['Current Status']] === 'Requisition Submitted for Approval')
//     .map(r => {
//       const reqRow = r.data;
//       const prID = reqRow[reqHeaders['Requisition ID']];
//       if (!prID) return null;

//       const itemsForReq = itemsData
//         .filter(itemRow => itemRow[itemHeaders['Requisition ID']] === prID)
//         .map(itemRow => ({
//           description: itemRow[itemHeaders['Item Description']],
//           quantity: itemRow[itemHeaders['Quantity']],
//           uom: itemRow[itemHeaders['UOM']],
//           unitPrice: itemRow[itemHeaders['Unit Price']],
//           totalPrice: itemRow[itemHeaders['Total Price']],
//         }));

//       const vendorId = reqRow[reqHeaders['Vendor ID']];
//       const vendor = getVendorDetails(vendorId); // Fetch vendor details

//       return {
//         id: prID,
//         date: reqRow[reqHeaders['Date of Requisition']],
//         expectedDeliveryDate: reqRow[reqHeaders['Expected Delivery Date']],
//         status: reqRow[reqHeaders['Current Status']],
//         totalValue: reqRow[reqHeaders['Total Value Incl. GST']],
//         purchaseCategory: reqRow[reqHeaders['Purchase Category']],
//         requestedBy: reqRow[reqHeaders['Requested By']],
//         site: reqRow[reqHeaders['Site']],
//         vendor: vendor,
//         items: itemsForReq,
//         rowIndex: r.rowIndex
//       };
//     })
//     .filter(req => req != null);

//   return { pendingRequisitions: JSON.stringify(pendingRequisitions) };
// }

// function getPendingRequisitions() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const reqSheet = ss.getSheetByName('Requisitions');
//   const itemsSheet = ss.getSheetByName('Items');

//   if (!reqSheet || !itemsSheet) {
//     throw new Error("Required sheets ('Requisitions', 'Items') not found.");
//   }

//   const reqHeaders = getHeaderMap(reqSheet);
//   const itemHeaders = getHeaderMap(itemsSheet);

//   const reqData = reqSheet.getRange(2, 1, reqSheet.getLastRow() - 1, reqSheet.getLastColumn()).getValues();
//   const itemsData = itemsSheet.getRange(2, 1, itemsSheet.getLastRow() - 1, itemsSheet.getLastColumn()).getValues();

//   const pendingRequisitions = reqData
//     .map((row, index) => ({ data: row, rowIndex: index + 2 }))
//     .filter(r => r.data[reqHeaders['Current Status']] === 'Requisition Submitted for Approval')
//     .map(r => {
//       const reqRow = r.data;
//       const prID = reqRow[reqHeaders['Requisition ID']];
//       if (!prID) return null;

//       const itemsForReq = itemsData
//         .filter(itemRow => itemRow[itemHeaders['Requisition ID']] === prID)
//         .map(itemRow => {
//           const quantity = parseFloat(itemRow[itemHeaders['Quantity Required']]) || 0;
//           const rate = parseFloat(itemRow[itemHeaders['Rate']]) || 0;
//           const gst = parseFloat(itemRow[itemHeaders['GST']]) || 0;
//           const subtotal = quantity * rate;
//           const totalValue = subtotal + (subtotal * gst / 100);
//           return {
//             itemId: itemRow[itemHeaders['Item ID']],
//             itemName: itemRow[itemHeaders['Item Name']],
//             purpose: itemRow[itemHeaders['Purpose / Application']],
//             quantity: quantity,
//             uom: itemRow[itemHeaders['UOM']],
//             rate: rate,
//             gst: gst,
//             warranty: itemRow[itemHeaders['Warranty, AMC']],
//             subtotal: subtotal,
//             totalValue: totalValue
//           };
//         });

//       const vendorId = reqRow[reqHeaders['Vendor ID']];
//       const vendorDetails = getVendorDetails(vendorId); 

//       const vendor = vendorDetails ? {
//         companyName: vendorDetails['COMPANY NAME'],
//         contactPerson: vendorDetails['CONTACT PERSON'],
//         contactNumber: vendorDetails['CONTACT NUMBER'],
//         email: vendorDetails['EMAIL ID']
//       } : {};

//       const totalSubtotal = itemsForReq.reduce((sum, item) => sum + (parseFloat(item.subtotal) || 0), 0);
//       const totalGST = itemsForReq.reduce((sum, item) => {
//         const itemGST = (parseFloat(item.subtotal) || 0) * (parseFloat(item.gst) || 0) / 100;
//         return sum + itemGST;
//       }, 0);
//       const grandTotal = totalSubtotal + totalGST;

//       return {
//         id: prID,
//         date: reqRow[reqHeaders['Date of Requisition']],
//         expectedDeliveryDate: reqRow[reqHeaders['Expected Delivery Date']],
//         status: reqRow[reqHeaders['Current Status']],
//         totalValue: grandTotal,
//         totalSubtotal: totalSubtotal,
//         totalGST: totalGST,
//         purchaseCategory: reqRow[reqHeaders['Purchase Category']],
//         requestedBy: reqRow[reqHeaders['Requested By']],
//         site: reqRow[reqHeaders['Site']],
//         vendor: vendor,
//         items: itemsForReq,
//         rowIndex: r.rowIndex
//       };
//     })
//     .filter(req => req != null);

//   return { pendingRequisitions: JSON.stringify(pendingRequisitions) };
// }
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
    .map((row, index) => ({ data: row, rowIndex: index + 2 }))
    .filter(r => r.data[reqHeaders['Current Status']] === 'Requisition Submitted for Approval')
    .map(r => {
      const reqRow = r.data;
      const prID = reqRow[reqHeaders['Requisition ID']];
      if (!prID) return null;

      const itemsForReq = itemsData
        .filter(itemRow => itemRow[itemHeaders['Requisition ID']] === prID)
        .map(itemRow => ({
          itemId: itemRow[itemHeaders['Item ID']],
          itemName: itemRow[itemHeaders['Item Name']],
          purpose: itemRow[itemHeaders['Purpose / Application']],
          quantity: itemRow[itemHeaders['Quantity Required']],
          uom: itemRow[itemHeaders['UOM']],
          rate: itemRow[itemHeaders['Rate']],
          gst: itemRow[itemHeaders['GST']],
          warranty: itemRow[itemHeaders['Warranty, AMC']],
          totalValue: itemRow[itemHeaders['Total Value (Incl. GST)']]
        }));

      const vendorId = reqRow[reqHeaders['Vendor ID']];
      const vendorDetails = getVendorDetails(vendorId); 

      const vendor = vendorDetails ? {
        companyName: vendorDetails['COMPANY NAME'],
        contactPerson: vendorDetails['CONTACT PERSON'],
        contactNumber: vendorDetails['CONTACT NUMBER'],
        email: vendorDetails['EMAIL ID']
      } : {};

      const totalSubtotal = itemsForReq.reduce((sum, item) => sum + ((parseFloat(item.quantity) || 0) * (parseFloat(item.rate) || 0)), 0);
      const totalGST = itemsForReq.reduce((sum, item) => {
        const itemGST = ((parseFloat(item.quantity) || 0) * (parseFloat(item.rate) || 0)) * ((parseFloat(item.gst) || 0) / 100);
        return sum + itemGST;
      }, 0);
      const grandTotal = totalSubtotal + totalGST;

      return {
        id: prID,
        date: reqRow[reqHeaders['Date of Requisition']],
        expectedDeliveryDate: reqRow[reqHeaders['Expected Delivery Date']],
        status: reqRow[reqHeaders['Current Status']],
        totalValue: grandTotal,
        totalSubtotal: totalSubtotal,
        totalGST: totalGST,
        purchaseCategory: reqRow[reqHeaders['Purchase Category']],
        requestedBy: reqRow[reqHeaders['Requested By']],
        site: reqRow[reqHeaders['Site']],
        vendor: vendor,
        items: itemsForReq,
        rowIndex: r.rowIndex
      };
    })
    .filter(req => req != null);

  return { pendingRequisitions: JSON.stringify(pendingRequisitions) };
}


// Function to get data for the PO approval page
function getPendingPOs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reqSheet = ss.getSheetByName('Requisitions');
  const poSheet = ss.getSheetByName('PO_Master');
  const itemsSheet = ss.getSheetByName('Items');

  if (!reqSheet || !poSheet || !itemsSheet) {
    return { pendingPOs: '[]' };
  }

  const reqHeaders = getHeaderMap(reqSheet);
  const poHeaders = getHeaderMap(poSheet);
  const itemHeaders = getHeaderMap(itemsSheet);

  const reqData = reqSheet.getRange(2, 1, reqSheet.getLastRow() - 1, reqSheet.getLastColumn()).getValues();
  const poData = poSheet.getRange(2, 1, poSheet.getLastRow() - 1, poSheet.getLastColumn()).getValues();
  const itemsData = itemsSheet.getRange(2, 1, itemsSheet.getLastRow() - 1, itemsSheet.getLastColumn()).getValues();

  const poMap = {};
  poData.forEach(row => {
    const reqId = row[poHeaders['Requisition ID']];
    poMap[reqId] = {
      id: row[poHeaders['PO No.']],
      date: row[poHeaders['PO Date']],
      totalValue: row[poHeaders['PO Amount']],
      attachPO: row[poHeaders['Attach PO']]
    };
  });

  const pendingPOs = reqData
    .filter(row => row[reqHeaders['Current Status']] === 'PO Submitted')
    .map(reqRow => {
      const prID = reqRow[reqHeaders['Requisition ID']];
      const poInfo = poMap[prID];
      if (!poInfo) return null;

      const itemsForReq = itemsData
        .filter(itemRow => itemRow[itemHeaders['Requisition ID']] === prID)
        .map(itemRow => ({
          description: itemRow[itemHeaders['Item Description']],
          quantity: itemRow[itemHeaders['Quantity']],
          uom: itemRow[itemHeaders['UOM']],
          unitPrice: itemRow[itemHeaders['Rate']],
          totalPrice: itemRow[itemHeaders['Total Value (Incl. GST)']],
        }));
      
      const vendorId = reqRow[reqHeaders['Vendor ID']];
      const vendor = getVendorDetails(vendorId); // Fetch vendor details

      return {
        poId: poInfo.id,
        prId: prID,
        date: reqRow[reqHeaders['Date of Requisition']],
        totalValue: reqRow[reqHeaders['Total Value Incl. GST']],
        requestedBy: reqRow[reqHeaders['Requested By']],
        site: reqRow[reqHeaders['Site']],
        vendor: vendor, // Attach vendor object
        items: itemsForReq,
        poDetails: poInfo
      };
    })
    .filter(po => po !== null);

  return { pendingPOs: JSON.stringify(pendingPOs) };
}

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
      let approvalStatus = '';
      if (action === 'Approved') {
        approvalStatus = `Approved - ${userEmail} - ${timestamp}`;
        updateRequisitionStatus(prID, 'Requisition Approved');
      } else {
        approvalStatus = `Declined - ${userEmail} - ${timestamp}`;
        updateRequisitionStatus(prID, 'Requisition Rejected');
      }
      reqSheet.getRange(i + 2, headers['Approval Status'] + 1).setValue(approvalStatus);
      reqSheet.getRange(i + 2, headers['Approver Remarks'] + 1).setValue(remarks);
      return { success: true, message: `Requisition ${prID} has been ${action}.` };
    }
  }
  throw new Error(`Requisition with ID ${prID} not found.`);
}

function processPOApproval(poID, action, remarks) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userEmail = Session.getActiveUser().getEmail();
  const timestamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss");

  // Update PO_Master sheet
  const poSheet = ss.getSheetByName('PO_Master');
  const poHeaders = getHeaderMap(poSheet);
  const poData = poSheet.getRange(2, 1, poSheet.getLastRow() - 1, poSheet.getLastColumn()).getValues();

  let reqId = null;
  let poRowIndex = -1;

  for (let i = 0; i < poData.length; i++) {
    if (poData[i][poHeaders['PO No.']] === poID) {
      reqId = poData[i][poHeaders['Requisition ID']];
      poRowIndex = i + 2; // 1-based index for getRange
      break;
    }
  }

  if (!reqId) {
    throw new Error(`PO with ID ${poID} not found in PO_Master.`);
  }

  const poStatus = action === 'Approved' ? 'Approved' : 'Rejected';
  poSheet.getRange(poRowIndex, poHeaders['PO Status'] + 1).setValue(poStatus);
  poSheet.getRange(poRowIndex, poHeaders['PO Remarks'] + 1).setValue(remarks);

  // Update Requisitions sheet
  const reqSheet = ss.getSheetByName('Requisitions');
  const reqHeaders = getHeaderMap(reqSheet);
  const reqData = reqSheet.getRange(2, 1, reqSheet.getLastRow() - 1, reqSheet.getLastColumn()).getValues();

  for (let i = 0; i < reqData.length; i++) {
    if (reqData[i][reqHeaders['Requisition ID']] === reqId) {
      const newStatus = `PO ${action} - ${userEmail} - ${timestamp}`;
      reqSheet.getRange(i + 2, reqHeaders['Current Status'] + 1).setValue(newStatus);
      
      // Also update the approver remarks in the main requisition
      const approvalMessage = `PO ${action} by ${userEmail} on ${timestamp}. Remarks: ${remarks}`;
      reqSheet.getRange(i + 2, reqHeaders['Approver Remarks'] + 1).setValue(approvalMessage);
      
      return { success: true, message: `PO ${poID} has been ${action}.` };
    }
  }

  throw new Error(`Requisition with ID ${reqId} (associated with PO ${poID}) not found.`);
}

function getDashboardStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reqSheet = ss.getSheetByName('Requisitions');
  const poSheet = ss.getSheetByName('PO_Master');

  if (!reqSheet) {
    return {
      totalRequisitions: 0,
      pendingPRApproval: 0,
      pendingPOApproval: 0,
      approvedPR: 0,
      approvedPO: 0,
      recentRequisitions: []
    };
  }

  const reqHeaders = getHeaderMap(reqSheet);
  const reqData = reqSheet.getRange(2, 1, reqSheet.getLastRow() - 1, reqSheet.getLastColumn()).getValues();

  let pendingPRApproval = 0;
  let pendingPOApproval = 0;
  let approvedPR = 0;

  reqData.forEach(row => {
    const status = row[reqHeaders['Current Status']];
    if (status === 'Requisition Submitted for Approval') {
      pendingPRApproval++;
    } else if (status === 'PO Submitted') {
      pendingPOApproval++;
    } else if (status === 'Requisition Approved') {
      approvedPR++;
    }
  });

  let approvedPO = 0;
  if (poSheet) {
    const poHeaders = getHeaderMap(poSheet);
    const poData = poSheet.getRange(2, 1, poSheet.getLastRow() - 1, poSheet.getLastColumn()).getValues();
    poData.forEach(row => {
      if (row[poHeaders['PO Status']] === 'Approved') {
        approvedPO++;
      }
    });
  }

  const recentRequisitions = reqData.slice(-5).reverse().map(row => {
    const vendorId = row[reqHeaders['Vendor ID']];
    const vendor = getVendorDetails(vendorId);
    return {
      id: row[reqHeaders['Requisition ID']],
      items: 'Details unavailable', // This would require reading the 'Items' sheet and is complex for a summary
      vendor: vendor ? vendor.companyName : 'N/A',
      amount: row[reqHeaders['Total Value Incl. GST']],
      status: row[reqHeaders['Current Status']],
      date: new Date(row[reqHeaders['Date of Requisition']]).toLocaleDateString()
    };
  });

  return {
    totalRequisitions: reqData.length,
    pendingPRApproval: pendingPRApproval,
    pendingPOApproval: pendingPOApproval,
    approvedPR: approvedPR,
    approvedPO: approvedPO,
    recentRequisitions: recentRequisitions
  };
}

/**
 * Fetches Requisition IDs that are approved and ready for PO creation.
 * @returns {string[]} A list of approved Requisition IDs.
 */
function getApprovedRequisitions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reqSheet = ss.getSheetByName('Requisitions');
  if (!reqSheet) return [];
  
  const headers = getHeaderMap(reqSheet);
  const data = reqSheet.getRange(2, 1, reqSheet.getLastRow() - 1, reqSheet.getLastColumn()).getValues();
  
  const approvedRequisitions = data
    .filter(row => row[headers['Current Status']] === 'Requisition Approved') 
    .map(row => row[headers['Requisition ID']]);
    
  return approvedRequisitions;
}
/**
 * Fetches details for a given Vendor ID from the Vendor_Master sheet.
 * @param {string} vendorId The Vendor ID.
 * @returns {object} An object containing vendor details.
 */
function getVendorDetails(vendorId) {
  if (!vendorId) return null;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const vendorSheet = ss.getSheetByName('Vendor_Master');
  if (!vendorSheet) {
    Logger.log("Sheet 'Vendor_Master' not found.");
    return null;
  }

  const headers = getHeaderMap(vendorSheet);
  const data = vendorSheet.getRange(2, 1, vendorSheet.getLastRow() - 1, vendorSheet.getLastColumn()).getValues();

  for (const row of data) {
    if (row[headers['VENDOR ID']] === vendorId) {
      const vendorDetails = {};
      Object.keys(headers).forEach(header => {
        vendorDetails[header] = row[headers[header]];
      });
      return vendorDetails;
    }
  }
  return null;
}


/**
 * Fetches all details for a given Requisition ID to populate the PO form.
 * @param {string} prID The Requisition ID.
 * @returns {object} An object containing requisition details and its line items.
 */
// function getRequisitionDetailsForPO(prID) {
//   if (!prID) {
//     throw new Error("Requisition ID is required.");
//   }
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const reqSheet = ss.getSheetByName('Requisitions');
//   const itemsSheet = ss.getSheetByName('Items');

//   if (!reqSheet || !itemsSheet) {
//     throw new Error("Required sheets ('Requisitions', 'Items') not found.");
//   }

//   const reqHeaders = getHeaderMap(reqSheet);
//   const itemHeaders = getHeaderMap(itemsSheet);
//   const reqData = reqSheet.getRange(2, 1, reqSheet.getLastRow() - 1, reqSheet.getLastColumn()).getValues();
//   const itemsData = itemsSheet.getRange(2, 1, itemsSheet.getLastRow() - 1, itemsSheet.getLastColumn()).getValues();

//   let requisitionDetails = {};
//   let found = false;

//   for (const row of reqData) {
//     if (row[reqHeaders['Requisition ID']] === prID) {
//       requisitionDetails = {
//         requisitionId: row[reqHeaders['Requisition ID']],
//         date: row[reqHeaders['Date of Requisition']],
//         requestedBy: row[reqHeaders['Requested By']],
//         site: row[reqHeaders['Site']],
//         deliveryLocation: row[reqHeaders['Delivery Location']],
//         purchaseCategory: row[reqHeaders['Purchase Category']],
//         totalValue: row[reqHeaders['Total Value Incl. GST']],
//         paymentTerms: row[reqHeaders['Payment Terms']],
//         deliveryTerms: row[reqHeaders['Delivery Terms']],
//         expectedDeliveryDate: row[reqHeaders['Expected Delivery Date']],
//       };

//       const vendorId = row[reqHeaders['Vendor ID']];
//       if (vendorId) {
//         requisitionDetails.vendor = getVendorDetails(vendorId);
//       } else {
//         requisitionDetails.vendor = {};
//       }

//       found = true;
//       break;
//     }
//   }

//   if (!found) {
//     throw new Error(`Requisition ${prID} not found.`);
//   }

//   const itemsForReq = itemsData
//     .filter(itemRow => itemRow[itemHeaders['Requisition ID']] === prID)
//     .map(itemRow => ({
//       itemName: itemRow[itemHeaders['Item Name']],
//       purpose: itemRow[itemHeaders['Purpose / Application']],
//       quantity: itemRow[itemHeaders['Quantity Required']],
//       uom: itemRow[itemHeaders['UOM']],
//       rate: itemRow[itemHeaders['Rate']],
//       gst: itemRow[itemHeaders['GST']],
//       totalCost: itemRow[itemHeaders['Total Value (Incl. GST)']] 
//     }));
  
//   requisitionDetails.items = itemsForReq;
//   return requisitionDetails;
// }

// Handle PO form submission
// function getRequisitionDetailsForPO(prID) {
//   if (!prID) return null;
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const reqSheet = ss.getSheetByName('Requisitions');
//   const itemsSheet = ss.getSheetByName('Items');
//   if (!reqSheet || !itemsSheet) return null;

//   const reqHeaders = getHeaderMap(reqSheet);
//   const itemHeaders = getHeaderMap(itemsSheet);
//   const reqData = reqSheet.getRange(2, 1, reqSheet.getLastRow() - 1, reqSheet.getLastColumn()).getValues();
//   const itemsData = itemsSheet.getRange(2, 1, itemsSheet.getLastRow() - 1, itemsSheet.getLastColumn()).getValues();

//   let requisitionDetails = {};
//   let found = false;

//   for (const row of reqData) {
//     if (row[reqHeaders['Requisition ID']] === prID) {
//       requisitionDetails = {
//         paymentTerms: row[reqHeaders['Payment Terms']],
//         deliveryTerms: row[reqHeaders['Delivery Terms']],
//         expectedDeliveryDate: row[reqHeaders['Expected Delivery Date']],
//         site: row[reqHeaders['Site']],
//         vendorRegistered: row[reqHeaders['Is Vendor Registered']],
//         vendor: getVendorDetails(row[reqHeaders['Vendor ID']])
//       };
//       found = true;
//       break;
//     }
//   }
//   if (!found) return null;

//   const itemsForReq = itemsData
//     .filter(itemRow => itemRow[itemHeaders['Requisition ID']] === prID)
//     .map(itemRow => ({
//       itemName: itemRow[itemHeaders['Item Name']],
//       purpose: itemRow[itemHeaders['Purpose / Application']],
//       quantity: itemRow[itemHeaders['Quantity Required']],
//       uom: itemRow[itemHeaders['UOM']],
//       rate: itemRow[itemHeaders['Rate']],
//       gst: itemRow[itemHeaders['GST']],
//       totalCost: itemRow[itemHeaders['Total Value (Incl. GST)']]
//     }));

//   requisitionDetails.items = itemsForReq;
//   return requisitionDetails;
// }

// Replace the original `getRequisitionDetailsForPO` function
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
      requisitionDetails = {
        requisitionId: row[reqHeaders['Requisition ID']],
        date: row[reqHeaders['Date of Requisition']],
        requestedBy: row[reqHeaders['Requested By']],
        site: row[reqHeaders['Site']],
        deliveryLocation: row[reqHeaders['Delivery Location']],
        purchaseCategory: row[reqHeaders['Purchase Category']],
        totalValue: row[reqHeaders['Total Value Incl. GST']],
        paymentTerms: row[reqHeaders['Payment Terms']],
        deliveryTerms: row[reqHeaders['Delivery Terms']],
        expectedDeliveryDate: row[reqHeaders['Expected Delivery Date']],
        isVendorRegistered: row[reqHeaders['Is the Vendor Registered with Us?']],
        vendorGSTCertificate: row[reqHeaders['Vendor GST Certificate']],
        vendorPANCard: row[reqHeaders['Vendor PAN Card']],
        vendorCancelledCheque: row[reqHeaders['Cancelled Cheque']]
      };

      const vendorId = row[reqHeaders['Vendor ID']];
      if (vendorId) {
        requisitionDetails.vendor = getVendorDetails(vendorId);
        requisitionDetails.vendorId = vendorId;
      } else {
        requisitionDetails.vendor = {};
        requisitionDetails.vendorId = '';
      }
      
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
// function submitPO(formData, lineItems) {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const poSheet = ss.getSheetByName('PO_Master'); 
//   const poHeaders = getHeaderMap(poSheet);

//   if (!poSheet) {
//     throw new Error("Sheet 'PO_Master' not found. Please check the sheet name in your Google Spreadsheet.");
//   }

//   // Use the PO Number provided by the user
//   const poNumber = formData.poNumber;
//   const vendorId = getVendorIdFromRequisition(formData.reqId);

//   const totalValue = lineItems.reduce((sum, item) => sum + parseFloat(item.totalCost || 0), 0);

//   // Create a new row array based on the provided column order
//   const newRow = new Array(poSheet.getLastColumn()).fill('');
//   newRow[poHeaders['Requisition ID']] = formData.reqId;
//   newRow[poHeaders['PO No.']] = poNumber;
//   newRow[poHeaders['Vendor ID']] = vendorId; // Adding Vendor ID
//   newRow[poHeaders['PO Date']] = formData.poDate;
//   newRow[poHeaders['PO Amount']] = totalValue.toFixed(2);
//   newRow[poHeaders['Attach PO']] = formData.attachPO;
//   newRow[poHeaders['PO Prepared By']] = formData.poPreparedBy;
//   newRow[poHeaders['PO Status']] = 'Submitted';
//   newRow[poHeaders['PO Remarks']] = '';
  
//   poSheet.appendRow(newRow);

//   // Update the original requisition's status
//   updateRequisitionStatus(formData.reqId, 'PO Submitted');

//   return { success: true, poID: poNumber };
// }

// Helper function to get Vendor ID from Requisitions sheet
// Replace the original `submitPO` function
function submitPO(formData, lineItems) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const poSheet = ss.getSheetByName('PO_Master'); 
  const poHeaders = getHeaderMap(poSheet);

  if (!poSheet) {
    throw new Error("Sheet 'PO_Master' not found. Please check the sheet name in your Google Spreadsheet.");
  }

  const poNumber = formData.poNumber;
  const vendorId = getVendorIdFromRequisition(formData.reqId);

  const totalValue = lineItems.reduce((sum, item) => sum + parseFloat(item.totalCost || 0), 0);

  const newRow = new Array(poSheet.getLastColumn()).fill('');
  newRow[poHeaders['Requisition ID']] = formData.reqId;
  newRow[poHeaders['PO No.']] = poNumber;
  newRow[poHeaders['Vendor ID']] = vendorId;
  newRow[poHeaders['PO Date']] = formData.poDate;
  newRow[poHeaders['PO Amount']] = totalValue.toFixed(2);
  newRow[poHeaders['Attach PO']] = formData.attachPO;
  newRow[poHeaders['PO Prepared By']] = formData.poPreparedBy;
  newRow[poHeaders['PO Status']] = 'Submitted';
  newRow[poHeaders['PO Remarks']] = '';
  
  poSheet.appendRow(newRow);

  updateRequisitionStatus(formData.reqId, 'PO Submitted');

  return { success: true, poID: poNumber };
}
// function getVendorIdFromRequisition(prID) {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const reqSheet = ss.getSheetByName('Requisitions');
//   const headers = getHeaderMap(reqSheet);
//   const data = reqSheet.getRange(2, 1, reqSheet.getLastRow() - 1, reqSheet.getLastColumn()).getValues();
//   for (let i = 0; i < data.length; i++) {
//     if (data[i][headers['Requisition ID']] === prID) {
//       return data[i][headers['Vendor ID']];
//     }
//   }
//   return null;
// }

// Add this new function
function getVendorIdFromRequisition(prID) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reqSheet = ss.getSheetByName('Requisitions');
  const headers = getHeaderMap(reqSheet);
  const data = reqSheet.getRange(2, 1, reqSheet.getLastRow() - 1, reqSheet.getLastColumn()).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][headers['Requisition ID']] === prID) {
      return data[i][headers['Vendor ID']];
    }
  }
  return null;
}

// Helper: Get next PR serial number for site and month/year
function getNextPRSerial(site) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Requisitions');
  const lastRow = sheet.getLastRow();
  
  // Get current Month Name and Year
  const now = new Date();
  const monthNames = [
    "Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
  ];
  const month = monthNames[now.getMonth()];
  const year = now.getFullYear();
  const monthYear = `${month}${year}`;

  if (lastRow < 2) {
    return { serial: '001', monthYear: monthYear };
  }

  // Pattern: PR-<Site>-<MonthYear>/<Serial>
  const pattern = new RegExp(`PR-${site}-${monthYear}/(\\d{3})`);
  const allIds = sheet.getRange(2, 3, lastRow - 1, 1).getValues().flat(); // Column 3: Requisition ID

  const serial = (
    Math.max(
      0, // Start with 0 in case no previous IDs match
      ...allIds
        .map(id => {
          const match = (id || '').toString().match(pattern);
          return match ? parseInt(match[1], 10) : 0;
        })
    ) + 1
  ).toString().padStart(3, '0');

  return { serial, monthYear };
}

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

  let vendorId = '';
  if (formData.isVendorRegistered === 'Yes') {
    vendorId = formData.vendorId;
  } else {
    // Create a new vendor and get the new ID
    const newVendor = {
      companyName: formData.companyName,
      contactPerson: formData.contactPerson,
      contactNumber: formData.contactNumber,
      emailId: formData.emailId,
      providingSites: formData.providingSites.join(', '),
      bankName: formData.bankName,
      accHolderName: formData.accHolderName,
      accNumber: formData.accNumber,
      branchName: formData.branchName,
      ifscCode: formData.ifscCode,
      gstNumber: formData.gstNumber,
      vendorPanNo: formData.vendorPanNo,
      vendorAddress: formData.vendorAddress,
      vendorGSTCertificate: formData.vendorGSTCertificate,
      vendorPANCard: formData.vendorPANCard,
      cancelledCheque: formData.cancelledCheque
    };
    vendorId = addVendorToMaster(newVendor);
  }

  // Save master details in the correct column order based on your CSV file
  const masterRow = [
    new Date(), // Timestamp
    formData.dateOfRequisition, // Date of Requisition
    prID, // Requisition ID
    'Pending Approval', // Approval Status
    'Requisition Submitted for Approval', // Current Status
    '', // Approver Remarks
    totalValue.toFixed(2), // Total Value Incl. GST
    formData.purchaseCategory, // Purchase Category
    formData.paymentTerms, // Payment Terms
    formData.deliveryTerms, // Delivery Terms
    formData.requestedBy, // Requested By
    formData.site, // Site
    formData.deliveryLocation, // Delivery Location
    formData.isVendorRegistered, // Is the Vendor Registered with Us?
    vendorId, // Vendor ID
    formData.quotationPI, // Upload Quotation / PI Final Agreed
    formData.supportingDocs, // Upload Supporting Documents
    formData.isCustomerReimbursable, // Is the customer reimbursable?
    '', // Approved PR Link
    '', // PDF Link
    '', // Approval Link(View)
    Session.getActiveUser().getEmail(), // Email Address
    formData.expectedDeliveryDate // Expected Delivery Date
  ];
  reqSheet.appendRow(masterRow);

  // Save line items
  lineItems.forEach((item, index) => {
    const itemID = prID + '-' + (index + 1).toString().padStart(2, '0');
    itemsSheet.appendRow([
      prID,
      itemID,
      item.itemName,
      item.purpose,
      item.quantity,
      item.uom,
      item.rate,
      item.gst,
      item.warranty,
      item.totalValue
    ]);
  });

  return { success: true, prID: prID };
}

// Handle file uploads (returns Drive link)
function uploadFile(blob, folderName) {
  const folder = DriveApp.getFoldersByName(folderName).hasNext()
    ? DriveApp.getFoldersByName(folderName).next()
    : DriveApp.createFolder(folderName);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function getAllVendors() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Vendor_Master');
  if (!sheet) {
    return [];
  }
  const data = sheet.getDataRange().getValues();``
  const headers = data[0];
  const headerMap = {};
  headers.forEach((header, i) => {
    headerMap[header] = i;
  });

  return data.slice(1).map(row => {
    const vendor = {};
    headers.forEach((header, i) => {
      vendor[header] = row[i];
    });
    return vendor;
  });
}

/**
 * Add a new vendor to Vendor_Master sheet.
 * @param {Object} vendor Vendor details.
 * @returns {string} Vendor ID.
 */
function addVendorToMaster(vendor) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Vendor_Master');
  if (!sheet) throw new Error('Vendor_Master sheet not found');
  const lastRow = sheet.getLastRow();
  // Find last serial number
  let serial = 1;
  if (lastRow > 1) {
    const ids = sheet.getRange(2, 1, lastRow - 1).getValues().map(r => r[0]);
    const serials = ids.map(id => {
      const match = id && id.match(/V\/\d{2}\/CRPL\/(\d{3})/);
      return match ? parseInt(match[1], 10) : 0;
    });
    serial = Math.max(...serials, 0) + 1;
  }
  const year = (new Date().getFullYear() % 100).toString().padStart(2, '0');
  const vendorId = `V/${year}/CRPL/${serial.toString().padStart(3, '0')}`;
  // Get providing sites as comma-separated string
  let providingSites = vendor.providingSites;
  if (Array.isArray(providingSites)) providingSites = providingSites.join(', ');
  // Prepare row (no GST Certificate, add Providing Sites)
  const row = [
    vendorId,
    vendor.companyName,
    vendor.contactPerson,
    vendor.contactNumber,
    vendor.emailId,
    vendor.bankName,
    vendor.accHolderName,
    vendor.accNumber,
    vendor.branchName,
    vendor.ifscCode,
    vendor.gstNumber,
    providingSites,
    vendor.vendorPanNo,
    vendor.vendorAddress,
    vendor.vendorGSTCertificate,
    vendor.vendorPANCard,
    vendor.cancelledCheque,
    Session.getActiveUser().getEmail()
  ];
  sheet.appendRow(row);
  return vendorId;
}