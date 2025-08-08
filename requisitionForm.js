function requisitionFormPage(user) {
    // This function returns the HTML for the requisition form.
    // I have reconstructed the form based on the fields in your error messages.
    // Please adjust if your actual form fields are different.
    return `
    <!DOCTYPE html>
<html lang="en">
<head>
    <base target="_top">
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Purchase Requisition Form</title>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0-beta3/css/all.min.css">
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700&display=swap');
        :root {
            --primary-color: #2962ff;
            --background-color: #f4f7fc;
            --card-background: #ffffff;
            --text-color: #333;
            --subtle-text-color: #6c757d;
            --border-color: #e0e0e0;
            --shadow: 0 4px 8px rgba(0, 0, 0, 0.05);
        }
        * { margin: 0; padding: 0; box-sizing: border-box; font-family: 'Poppins', sans-serif; }
        body { background-color: var(--background-color); color: var(--text-color); }
        .container { width: 100%; max-width: 1400px; margin: 0 auto; padding: 0 20px; }
        .header-wrapper { background-color: var(--card-background); padding: 10px 0; box-shadow: var(--shadow); }
        .header { display: flex; justify-content: space-between; align-items: center; }
        .header-title { display: flex; align-items: center; font-size: 1.5rem; font-weight: 600; color: var(--primary-color); }
        .header-title i { margin-right: 10px; }
        .user-info { display: flex; align-items: center; gap: 15px; }
        .user-info span { font-weight: 500; }
        .user-info .employee-btn {
            padding: 8px 12px; background-color: var(--primary-color); color: white;
            border: none; border-radius: 5px; cursor: pointer; display: flex; align-items: center; gap: 5px;
        }
        .nav-bar-wrapper { background-color: var(--primary-color); }
        .nav-bar { display: flex; align-items: center; }
        .nav-bar a {
            color: white; text-decoration: none; padding: 15px 20px; display: flex; align-items: center; gap: 8px;
            font-weight: 500; transition: background-color 0.3s;
        }
        .nav-bar a:hover { background-color: rgba(255, 255, 255, 0.1); }
        .nav-bar a.active { background-color: #1a53e0; }
        .main-content { padding-top: 20px; padding-bottom: 40px; }
        .form-container { background-color: var(--card-background); padding: 30px; border-radius: 8px; box-shadow: var(--shadow); }
        .form-container h2 {
            font-weight: 600; margin-bottom: 25px; color: var(--primary-color);
            border-bottom: 2px solid var(--border-color); padding-bottom: 15px;
        }
        .section { margin-bottom: 30px; border-bottom: 1px solid #eee; padding-bottom: 20px; }
        .section:last-child { border-bottom: none; }
        .section h3 { font-size: 1.1rem; font-weight: 500; margin-bottom: 20px; color: #555; }
        .form-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 20px; }
        .form-field { display: flex; flex-direction: column; }
        label { margin-bottom: 8px; font-weight: 500; font-size: 0.9rem; }
        input[type="text"], input[type="date"], input[type="email"], input[type="number"], select {
            width: 100%; padding: 10px; border: 1px solid var(--border-color); border-radius: 5px; font-size: 1rem;
        }
        input[type="file"] { padding: 8px; border: 1px solid var(--border-color); border-radius: 5px; }
        .line-items-table { width: 100%; border-collapse: collapse; margin-top: 10px; }
        .line-items-table th, .line-items-table td { padding: 12px; text-align: left; border-bottom: 1px solid var(--border-color); }
        .line-items-table th { background-color: #f8f9fa; font-weight: 500; }
        .line-items-table input[name="quantity"], .line-items-table input[name="gst"] { width: 70px; }
        .line-items-table select[name="uom"] { width: 120px; }
        .add-btn, .remove-btn, button[type="submit"] {
            padding: 10px 20px; border: none; border-radius: 5px; cursor: pointer; font-weight: 500; font-size: 0.9rem; transition: background-color 0.3s;
        }
        .add-btn { background-color: #e0e7ff; color: var(--primary-color); margin-top: 10px; }
        .add-btn:hover { background-color: #c7d2fe; }
        .remove-btn { background-color: #fee2e2; color: #ef4444; padding: 10px; line-height: 1; }
        .remove-btn:hover { background-color: #fecaca; }
        
        .submission-container {
            display: flex;
            flex-direction: column;
            align-items: flex-end;
            margin-top: 20px;
        }

        .totals-section {
            width: 100%;
            max-width: 350px;
            margin-bottom: 20px;
        }
        .total-row {
            display: flex;
            justify-content: space-between;
            padding: 8px 0;
            font-size: 1rem;
            border-bottom: 1px solid #f0f0f0;
        }
        .total-row:last-child {
            border-bottom: none;
        }
        .total-row span:first-child {
            color: var(--subtle-text-color);
        }
        .total-row span:last-child {
            font-weight: 600;
            color: var(--text-color);
        }
        .grand-total {
            font-size: 1.2rem;
            font-weight: bold;
            border-top: 2px solid var(--border-color);
            margin-top: 5px;
        }
        .grand-total span {
            color: var(--primary-color) !important;
        }

        button[type="submit"] { background-color: var(--primary-color); color: white; font-size: 1rem; width: 100%; max-width: 350px; }
        button[type="submit"]:hover { background-color: #1a53e0; }
        .hidden { display: none; }
        #msg { margin-top: 20px; font-weight: bold; }
        .success { color: #22c55e; }
        .error { color: #ef4444; }
        /* Loader styles */
        #loader {
          display: none;
          position: fixed;
          top: 0; left: 0; width: 100vw; height: 100vh;
          background: rgba(255,255,255,0.6);
          z-index: 9999;
          align-items: center;
          justify-content: center;
        }
        #loader .spinner {
          border: 6px solid #f3f3f3;
          border-top: 6px solid var(--primary-color);
          border-radius: 50%;
          width: 60px;
          height: 60px;
          animation: spin 1s linear infinite;
        }
        @keyframes spin {
          0% { transform: rotate(0deg); }
          100% { transform: rotate(360deg); }
        }
        #loader span {
          display: block;
          margin-top: 16px;
          color: var(--primary-color);
          font-weight: 600;
          font-size: 1.1rem;
          text-align: center;
        }

      /* Add your form styles here */
      .form-section { margin-bottom: 20px; border: 1px solid #ccc; padding: 15px; border-radius: 5px; }
      .form-section h3 { margin-top: 0; }
      .form-row { display: flex; gap: 20px; margin-bottom: 10px; }
      .form-row > div { flex: 1; }
      label { display: block; margin-bottom: 5px; font-weight: bold; }
      input[type="text"], input[type="date"], input[type="email"], select, textarea {
        width: 100%; padding: 8px; border-radius: 4px; border: 1px solid #ccc;
      }
      /* Hide sections by default */
      #unregisteredVendorDetails, #registeredVendorDetails { display: none; }
    </style>
</head>
<body>
    <div class="main-content container">
        <div class="form-container">
            <h2>Purchase Requisition Form</h2>
            <form id="prForm">
                <div class="section">
                    <h3>Master Details</h3>
                    <div class="form-grid">
                        <div class="form-field">
                            <label>Date of Requisition:
                                <input type="date" name="dateOfRequisition" id="dateOfRequisition" readonly required>
                            </label>
                        </div>
                        <div class="form-field">
                            <label>Requested By:
                                <input type="text" name="requestedBy" id="requestedBy" required>
                            </label>
                        </div>
                        <div class="form-field">
                            <label>Requester Email:
                                <input type="email" name="emailAddress" id="emailAddress" value="${user.email}" readonly>
                            </label>
                        </div>
                        <div class="form-field">
                            <label>Site:</label>
                            <select name="site" id="site" required>
                                <option value="">Select Site...</option>
                                <option>Bhubaneswar</option><option>Detroj</option><option>Dhulagarh</option><option>Kheda</option><option>Kolkata</option><option>Noida</option><option>Pune</option>
                            </select>
                        </div>
                        <div class="form-field">
                            <label>Purchase Category:</label>
                            <select name="purchaseCategory" required>
                                <option value="">Select Category...</option>
                                <option>Consumables</option>
                                <option>Maintenance Capex</option>
                                <option>Operations Capex</option>
                                <option>Project/Site Capex</option>
                                <option>Service</option>
                            </select>
                        </div>
                        <div class="form-field">
                            <label>Is the Vendor Registered?</label>
                            <select name="isVendorRegistered" id="isVendorRegistered" required onchange="toggleVendorFields()">
                                <option value="">Select...</option><option value="Yes">Yes</option><option value="No">No</option>
                            </select>
                        </div>
                        <div class="form-field">
                            <label>Is the customer reimbursable?</label>
                            <select name="isCustomerReimbursable" required>
                                <option value="">Select...</option>
                                <option value="Yes">Yes</option>
                                <option value="No">No</option>
                            </select>
                        </div>
                    </div>
                </div>
                <div class="section hidden" id="vendorDetailsSection">
                  <h3>Vendor Details</h3>
                  <div id="registeredVendorFields" class="form-grid hidden">
                    <div class="form-field">
                      <label>Select Registered Vendor:
                        <select id="vendorSelect" name="vendorCompanyName" onchange="fillVendorDetails()" required>
                          <option value="">Select Vendor...</option>
                        </select>
                      </label>
                    </div>
                    <div class="form-field"><label>Company Name: <input type="text" name="companyName" id="companyName" readonly></label></div>
                    <div class="form-field"><label>Contact Person: <input type="text" name="contactPerson" id="contactPerson" readonly></label></div>
                    <div class="form-field"><label>Contact Number: <input type="text" name="contactNumber" id="contactNumber" readonly></label></div>
                    <div class="form-field"><label>Email ID: <input type="email" name="emailId" id="emailId" readonly></label></div>
                    <div class="form-field"><label>Bank Name: <input type="text" name="bankName" id="bankName" readonly></label></div>
                    <div class="form-field"><label>Acc Holder Name: <input type="text" name="accHolderName" id="accHolderName" readonly></label></div>
                    <div class="form-field"><label>Acc Number: <input type="text" name="accNumber" id="accNumber" readonly></label></div>
                    <div class="form-field"><label>Branch Name: <input type="text" name="branchName" id="branchName" readonly></label></div>
                    <div class="form-field"><label>IFSC Code: <input type="text" name="ifscCode" id="ifscCode" readonly></label></div>
                    <div class="form-field"><label>GST Number: <input type="text" name="gstNumber" id="gstNumber" readonly></label></div>
                    <div class="form-field"><label>Providing Sites: <input type="text" name="providingSites" id="providingSites" readonly></label></div>
                    <div class="form-field"><label>Vendor PAN No: <input type="text" name="vendorPanNo" id="vendorPanNo" readonly></label></div>
                    <div class="form-field"><label>Vendor Address: <input type="text" name="vendorAddress" id="vendorAddress" readonly></label></div>
                  </div>
                  <div id="unregisteredVendorFields" class="form-grid hidden">
                    <div class="form-field"><label>Company Name: <input type="text" name="companyName" required></label></div>
                    <div class="form-field"><label>Contact Person: <input type="text" name="contactPerson" required></label></div>
                    <div class="form-field"><label>Contact Number: <input type="text" name="contactNumber" required></label></div>
                    <div class="form-field"><label>Email ID: <input type="email" name="emailId" required></label></div>
                    <div class="form-field"><label>Bank Name: <input type="text" name="bankName" required></label></div>
                    <div class="form-field"><label>Acc Holder Name: <input type="text" name="accHolderName" required></label></div>
                    <div class="form-field"><label>Acc Number: <input type="text" name="accNumber" required></label></div>
                    <div class="form-field"><label>Branch Name: <input type="text" name="branchName" required></label></div>
                    <div class="form-field"><label>IFSC Code: <input type="text" name="ifscCode" required></label></div>
                    <div class="form-field"><label>GST Number: <input type="text" name="gstNumber"></label></div>
                    <div class="form-field">
                      <label>Providing Sites:</label>
                      <div id="providingSitesCheckboxes">
                        <label><input type="checkbox" name="providingSites" value="Bhubaneswar"> Bhubaneswar</label>
                        <label><input type="checkbox" name="providingSites" value="Detroj"> Detroj</label>
                        <label><input type="checkbox" name="providingSites" value="Dhulagarh"> Dhulagarh</label>
                        <label><input type="checkbox" name="providingSites" value="Kheda"> Kheda</label>
                        <label><input type="checkbox" name="providingSites" value="Kolkata"> Kolkata</label>
                        <label><input type="checkbox" name="providingSites" value="Noida"> Noida</label>
                        <label><input type="checkbox" name="providingSites" value="Pune"> Pune</label>
                      </div>
                    </div>
                    <div class="form-field"><label>Vendor PAN No: <input type="text" name="vendorPanNo"></label></div>
                    <div class="form-field"><label>Vendor Address: <input type="text" name="vendorAddress"></label></div>
                  </div>
                </div>
                <div class="section">
                    <h3>Attachments & Delivery</h3>
                    <div class="form-grid">
                        <div class="form-field"><label>Vendor GST Certificate: <input type="file" name="vendorGSTCertificate"></label></div>
                        <div class="form-field"><label>Vendor PAN Card: <input type="file" name="vendorPANCard"></label></div>
                        <div class="form-field"><label>Cancelled Cheque: <input type="file" name="cancelledCheque"></label></div>
                        <div class="form-field"><label>Upload Quotation/Final Agreed PI: <input type="file" name="quotationPI" required></label></div>
                        <div class="form-field"><label>Supporting Docs: <input type="file" name="supportingDocs"></label></div>
                        <div class="form-field"><label>Payment Terms: <input type="text" name="paymentTerms" required></label></div>
                        <div class="form-field"><label>Delivery Terms: <input type="text" name="deliveryTerms" required></label></div>
                        <div class="form-field"><label>Expected Delivery Date: <input type="date" name="expectedDeliveryDate" required></label></div>
                        <div class="form-field"><label>Delivery Location: <input type="text" name="deliveryLocation" required></label></div>
                    </div>
                </div>
                <div class="section">
                    <h3>Line Items</h3>
                    <table class="line-items-table" id="lineItemsTable">
                        <thead>
                            <tr>
                                <th>Item Name</th><th>Purpose</th><th>Qty</th><th>UOM</th><th>Rate</th><th>GST (%)</th><th>Warranty/AMC</th><th>Total</th><th>Action</th>
                            </tr>
                        </thead>
                        <tbody id="lineItemsBody"></tbody>
                    </table>
                    <button type="button" class="add-btn" onclick="addLineItem()"><i class="fas fa-plus"></i> Add Item</button>
                </div>
                
                <div class="submission-container">
                    <div class="totals-section">
                        <div class="total-row">
                            <span>Subtotal:</span>
                            <span id="subtotalDisplay">₹0.00</span>
                        </div>
                        <div class="total-row">
                            <span>GST Amount:</span>
                            <span id="gstDisplay">₹0.00</span>
                        </div>
                        <div class="total-row grand-total">
                            <span>Grand Total:</span>
                            <span id="grandTotalDisplay">₹0.00</span>
                        </div>
                    </div>
                </div>

                <div id="unregisteredVendorDetails">
                  <h4>New Vendor Details</h4>
                  <div class="form-row">
                    <div><label for="companyName">Company Name</label><input type="text" id="companyName" name="companyName"></div>
                    <div><label for="contactPerson">Contact Person</label><input type="text" id="contactPerson" name="contactPerson"></div>
                  </div>
                  <div class="form-row">
                    <div><label for="contactNumber">Contact Number</label><input type="text" id="contactNumber" name="contactNumber"></div>
                    <div><label for="emailId">Email ID</label><input type="email" id="emailId" name="emailId"></div>
                  </div>
                  <h4>Bank Details</h4>
                  <div class="form-row">
                    <div><label for="bankName">Bank Name</label><input type="text" id="bankName" name="bankName"></div>
                    <div><label for="accHolderName">Account Holder Name</label><input type="text" id="accHolderName" name="accHolderName"></div>
                  </div>
                  <div class="form-row">
                    <div><label for="accNumber">Account Number</label><input type="text" id="accNumber" name="accNumber"></div>
                    <div><label for="branchName">Branch Name</label><input type="text" id="branchName" name="branchName"></div>
                  </div>
                  <div class="form-row">
                    <div><label for="ifscCode">IFSC Code</label><input type="text" id="ifscCode" name="ifscCode"></div>
                  </div>
                </div>
              </div>
      
              <button type="submit">Submit Requisition</button>
            </form>
            <div id="msg"></div>
        </div>
    </div>
    <div id="loader"><div style="display:flex;flex-direction:column;align-items:center;"><div class="spinner"></div><span id="loader-text">Loading...</span></div></div>
    <script>
      let allVendorsData = []; // Store all vendor data globally

      document.addEventListener('DOMContentLoaded', function() {
        // Set today's date for the requisition date field
        const today = new Date();
        const yyyy = today.getFullYear();
        const mm = String(today.getMonth() + 1).padStart(2, '0'); // Months are 0-based
        const dd = String(today.getDate()).padStart(2, '0');
        document.getElementById('dateOfRequisition').value = \`\${yyyy}-\${mm}-\${dd}\`;
      
        // Add the first line item on page load
        addLineItem();
      });

      function toggleVendorFields() {
        const isRegistered = document.getElementById('isVendorRegistered').value === 'Yes';
        
        document.getElementById('vendorDetailsSection').classList.toggle('hidden', !isRegistered);
        document.getElementById('registeredVendorFields').classList.toggle('hidden', !isRegistered);
        document.getElementById('unregisteredVendorFields').classList.toggle('hidden', isRegistered);

        document.getElementById('vendorSelect').required = isRegistered;

        const unregisteredFields = document.querySelectorAll('#unregisteredVendorFields [required]');
        unregisteredFields.forEach(field => {
          field.required = !isRegistered;
        });
      }

      function fillVendorDetails() {
        const vendorId = document.getElementById('vendorSelect').value;
        const registeredFieldsContainer = document.getElementById('registeredVendorFields');

        if (!vendorId) {
          // Clear all fields if no vendor is selected
          const fields = registeredFieldsContainer.querySelectorAll('input');
          fields.forEach(field => field.value = '');
          return;
        }
        
        google.script.run.withSuccessHandler(vendor => {
          if (vendor) {
            // Map API response to form field IDs
            const fieldMapping = {
              'COMPANY NAME': 'companyName',
              'CONTACT PERSON': 'contactPerson',
              'CONTACT NUMBER': 'contactNumber',
              'EMAIL ID': 'emailId',
              'BANK NAME': 'bankName',
              'ACC HOLDER NAME': 'accHolderName',
              'ACC NUMBER': 'accNumber',
              'BRANCH NAME': 'branchName',
              'IFSC CODE': 'ifscCode',
              'GST NUMBER': 'gstNumber',
              'PROVIDING SITES': 'providingSites',
              'VENDOR PAN NO': 'vendorPanNo',
              'VENDOR ADDRESS': 'vendorAddress'
            };

            for (const [key, id] of Object.entries(fieldMapping)) {
              const element = document.getElementById(id);
              if (element) {
                element.value = vendor[key] || '';
              }
            }
          }
        }).getVendorDetails(vendorId);
      }

      // Populate vendor dropdown on page load
      google.script.run.withSuccessHandler(vendors => {
        allVendorsData = vendors; // Save for later use
        const vendorSelect = document.getElementById('vendorSelect');
        vendorSelect.innerHTML = '<option value="">Select a vendor</option>';
        vendors.forEach(vendor => {
          vendorSelect.innerHTML += '<option value="' + vendor['VENDOR ID'] + '">' + vendor['COMPANY NAME'] + '</option>';
        });
      }).getAllVendors();

      // Add your form submission logic here
      document.getElementById('prForm').addEventListener('submit', function(e) {
        e.preventDefault();
        
        const form = e.target;
        const submitBtn = document.querySelector('button[type="submit"]');
        submitBtn.disabled = true;

        showLoader('Submitting Requisition...');

        const formData = {
            dateOfRequisition: form.dateOfRequisition.value,
            site: form.site.value,
            purchaseCategory: form.purchaseCategory.value,
            isVendorRegistered: form.isVendorRegistered.value,
            isCustomerReimbursable: form.isCustomerReimbursable.value,
            paymentTerms: form.paymentTerms.value,
            deliveryTerms: form.deliveryTerms.value,
            expectedDeliveryDate: form.expectedDeliveryDate.value,
            deliveryLocation: form.deliveryLocation.value,
            requestedBy: form.requestedBy.value // Manually entered name
        };

        // Handle vendor details based on selection
        if (formData.isVendorRegistered === 'Yes') {
            formData.vendorId = document.getElementById('vendorSelect').value;
        } else {
            formData.companyName = document.querySelector('#unregisteredVendorFields input[name="companyName"]').value;
            formData.contactPerson = document.querySelector('#unregisteredVendorFields input[name="contactPerson"]').value;
            formData.contactNumber = document.querySelector('#unregisteredVendorFields input[name="contactNumber"]').value;
            formData.emailId = document.querySelector('#unregisteredVendorFields input[name="emailId"]').value;
            formData.bankName = document.querySelector('#unregisteredVendorFields input[name="bankName"]').value;
            formData.accHolderName = document.querySelector('#unregisteredVendorFields input[name="accHolderName"]').value;
            formData.accNumber = document.querySelector('#unregisteredVendorFields input[name="accNumber"]').value;
            formData.branchName = document.querySelector('#unregisteredVendorFields input[name="branchName"]').value;
            formData.ifscCode = document.querySelector('#unregisteredVendorFields input[name="ifscCode"]').value;
            formData.gstNumber = document.querySelector('#unregisteredVendorFields input[name="gstNumber"]').value;
            formData.vendorPanNo = document.querySelector('#unregisteredVendorFields input[name="vendorPanNo"]').value;
            formData.vendorAddress = document.querySelector('#unregisteredVendorFields input[name="vendorAddress"]').value;
            // Get selected providing sites from checkboxes
            const selectedSites = Array.from(document.querySelectorAll('#providingSitesCheckboxes input[name="providingSites"]:checked'))
                                     .map(cb => cb.value);
            formData.providingSites = selectedSites;
        }

        const lineItems = [];
        document.querySelectorAll('#lineItemsBody tr').forEach(row => {
            const item = {
                itemName: row.querySelector('[name="itemName"]').value,
                purpose: row.querySelector('[name="purpose"]').value,
                quantity: row.querySelector('[name="quantity"]').value,
                uom: row.querySelector('[name="uom"]').value,
                rate: row.querySelector('[name="rate"]').value,
                gst: row.querySelector('[name="gst"]').value,
                warranty: row.querySelector('[name="warranty"]').value,
                totalValue: row.querySelector('[name="totalValue"]').value
            };
            lineItems.push(item);
        });

        // Handle file uploads
        const fileUploads = {
            quotationPI: form.quotationPI,
            supportingDocs: form.supportingDocs
        };
        
        // This is a simplified example. A more robust solution would handle multiple files.
        const quotationFile = fileUploads.quotationPI.files[0];
        const supportingDocsFile = fileUploads.supportingDocs.files[0];

        const filePromises = [];
        if (quotationFile) {
            filePromises.push(toBase64(quotationFile).then(base64 => {
                return new Promise(resolve => {
                    google.script.run
                        .withSuccessHandler(url => resolve({ name: 'quotationPI', url: url }))
                        .withFailureHandler(() => resolve({ name: 'quotationPI', url: '' }))
                        .uploadFile({ name: quotationFile.name, mimeType: quotationFile.type, base64: base64 }, 'Quotation_Attachments');
                });
            }));
        }

        if (supportingDocsFile) {
             filePromises.push(toBase64(supportingDocsFile).then(base64 => {
                return new Promise(resolve => {
                    google.script.run
                        .withSuccessHandler(url => resolve({ name: 'supportingDocs', url: url }))
                        .withFailureHandler(() => resolve({ name: 'supportingDocs', url: '' }))
                        .uploadFile({ name: supportingDocsFile.name, mimeType: supportingDocsFile.type, base64: base64 }, 'Supporting_Documents');
                });
            }));
        }

        Promise.all(filePromises).then(results => {
            results.forEach(result => {
                if (result.name === 'quotationPI') formData.quotationPI = result.url;
                if (result.name === 'supportingDocs') formData.supportingDocs = result.url;
            });
            
            // Call the server-side function
            google.script.run
                .withSuccessHandler(resp => {
                    hideLoader();
                    const msgDiv = document.getElementById('msg');
                    if (resp && resp.success) {
                        msgDiv.textContent = 'Requisition submitted successfully! PR ID: ' + resp.prID;
                        msgDiv.className = 'success';
                        form.reset();
                        document.getElementById('lineItemsBody').innerHTML = ''; // Clear line items
                        addLineItem(); // Add one empty line item
                        updateGrandTotal();
                    } else {
                        msgDiv.textContent = 'Submission failed. Please try again.';
                        msgDiv.className = 'error';
                    }
                    submitBtn.disabled = false;
                })
                .withFailureHandler(err => {
                    hideLoader();
                    document.getElementById('msg').textContent = 'Submission failed: ' + err.message;
                    document.getElementById('msg').className = 'error';
                    submitBtn.disabled = false;
                })
                .submitPR(formData, lineItems);
        });
      });

      function addLineItem() {
          const tableBody = document.getElementById('lineItemsBody');
          const newRow = tableBody.insertRow();
          newRow.innerHTML = \`
              <td><input type="text" name="itemName" required></td>
              <td><input type="text" name="purpose"></td>
              <td><input type="number" name="quantity" min="1" oninput="calculateTotal(this)" required></td>
              <td><select name="uom">
                  <option value="">Select</option>
                  <option>KGS</option>
                  <option>PIECE</option>
                  <option>NOS</option>
                  <option>BOX</option>
                  <option>MTR</option>
                  <option>PKT</option>
                  <option>M3</option>
                  <option>SET</option>
                  <option>FT</option>
                  <option>LTR</option>
                  <option>MM</option>
                  <option>PAIRS</option>
                  <option>RMT</option>
                  <option>ROLL</option>
                  <option>SQF</option>
                  <option>SQM</option>
                  <option>TROLLY</option>
                  <option>GRAM</option>
                  <option>BOTTLE</option>
                  <option>BAGS</option>
              </select></td>
              <td><input type="number" name="rate" min="0" oninput="calculateTotal(this)" required></td>
              <td><input type="number" name="gst" min="0" max="100" value="18" oninput="calculateTotal(this)" required></td>
              <td><input type="text" name="warranty" placeholder="e.g., 1 year"></td>
              <td><input type="number" name="totalValue" readonly></td>
              <td><button type="button" class="remove-btn" onclick="removeLineItem(this)">×</button></td>
          \`;
      }

      function removeLineItem(button) {
          const row = button.parentNode.parentNode;
          row.parentNode.removeChild(row);
          updateGrandTotal();
      }

      function calculateTotal(input) {
          const row = input.closest('tr');
          const quantity = parseFloat(row.querySelector('[name="quantity"]').value) || 0;
          const rate = parseFloat(row.querySelector('[name="rate"]').value) || 0;
          const gst = parseFloat(row.querySelector('[name="gst"]').value) || 0;
          const totalInput = row.querySelector('[name="totalValue"]');
          
          if (quantity > 0 && rate > 0) {
              const subtotal = quantity * rate;
              const total = subtotal + (subtotal * (gst / 100));
              totalInput.value = total.toFixed(2);
          } else {
              totalInput.value = '';
          }
          updateGrandTotal();
      }

      function updateGrandTotal() {
          const allRows = document.querySelectorAll('#lineItemsBody tr');
          let grandSubtotal = 0;
          let grandGST = 0;
          let grandTotal = 0;

          allRows.forEach(row => {
              const quantity = parseFloat(row.querySelector('[name="quantity"]').value) || 0;
              const rate = parseFloat(row.querySelector('[name="rate"]').value) || 0;
              const gst = parseFloat(row.querySelector('[name="gst"]').value) || 0;
              
              if (quantity > 0 && rate > 0) {
                  const subtotal = quantity * rate;
                  const gstAmount = subtotal * (gst / 100);
                  const total = subtotal + gstAmount;

                  grandSubtotal += subtotal;
                  grandGST += gstAmount;
                  grandTotal += total;
              }
          });

          document.getElementById('subtotalDisplay').innerText = '₹' + grandSubtotal.toFixed(2);
          document.getElementById('gstDisplay').innerText = '₹' + grandGST.toFixed(2);
          document.getElementById('grandTotalDisplay').innerText = '₹' + grandTotal.toFixed(2);
      }
      
      // Helper functions for loader and file upload
      function showLoader(text) {
          document.getElementById('loader-text').textContent = text;
          document.getElementById('loader').style.display = 'flex';
      }

      function hideLoader() {
          document.getElementById('loader').style.display = 'none';
      }

      function toBase64(file) {
          return new Promise((resolve, reject) => {
              const reader = new FileReader();
              reader.onload = () => resolve(reader.result.split(',')[1]);
              reader.onerror = reject;
              reader.readAsDataURL(file);
          });
      }
    </script>
  `;
}