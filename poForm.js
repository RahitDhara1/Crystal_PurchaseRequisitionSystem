function poFormPage(user) {
  return `
<!DOCTYPE html>
<html lang="en">
<head>
  <base target="_top">
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Purchase Order Form</title>
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0-beta3/css/all.min.css">
  <style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700&display=swap');
    :root {
      --primary-color: #2962ff;
      --background-color: #f4f7fc;
      --card-background: #ffffff;
      --text-color: #333;
      --border-color: #e0e0e0;
      --shadow: 0 4px 8px rgba(0, 0, 0, 0.05);
    }
    * { margin: 0; padding: 0; box-sizing: border-box; font-family: 'Poppins', sans-serif; }
    body { background-color: var(--background-color); color: var(--text-color); }
    .container { width: 100%; max-width: 900px; margin: 0 auto; padding: 0 20px; }
    .form-container { background-color: var(--card-background); padding: 30px; border-radius: 8px; box-shadow: var(--shadow); margin-top: 40px; }
    h2 { font-weight: 600; margin-bottom: 25px; color: var(--primary-color); border-bottom: 2px solid var(--border-color); padding-bottom: 15px; }
    .section { margin-bottom: 30px; border-bottom: 1px solid #eee; padding-bottom: 20px; }
    .section:last-child { border-bottom: none; }
    .form-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 20px; }
    .form-field { display: flex; flex-direction: column; }
    label { margin-bottom: 8px; font-weight: 500; font-size: 0.9rem; }
    input[type="text"], input[type="date"], input[type="number"], select {
      width: 100%; padding: 10px; border: 1px solid var(--border-color); border-radius: 5px; font-size: 1rem;
    }
    input[readonly], select[readonly] { background-color: #f8f9fa; cursor: not-allowed; }
    input[type="file"] { padding: 8px; border: 1px solid var(--border-color); border-radius: 5px; }
    .line-items-table { width: 100%; border-collapse: collapse; margin-top: 10px; }
    .line-items-table th, .line-items-table td { padding: 12px; text-align: left; border-bottom: 1px solid var(--border-color); }
    .line-items-table th { background-color: #f8f9fa; font-weight: 500; }
    button[type="submit"] { background-color: var(--primary-color); color: white; font-size: 1rem; float: right; padding: 10px 20px; border: none; border-radius: 5px; cursor: pointer; }
    button[type="submit"]:hover { background-color: #1a53e0; }
    button:disabled { background-color: #a5b4fc; cursor: not-allowed; }
    #msg { margin-top: 20px; font-weight: bold; }
    .success { color: #22c55e; }
    .error { color: #ef4444; }
    #loader {
      display: none; position: fixed; top: 0; left: 0; width: 100vw; height: 100vh;
      background: rgba(255,255,255,0.6); z-index: 9999; align-items: center; justify-content: center;
    }
    #loader .spinner {
      border: 6px solid #f3f3f3; border-top: 6px solid var(--primary-color); border-radius: 50%;
      width: 60px; height: 60px; animation: spin 1s linear infinite;
    }
    @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
    #loader span { display: block; margin-top: 16px; color: var(--primary-color); font-weight: 600; font-size: 1.1rem; text-align: center; }
  </style>
</head>
<body>
  <div class="main-content container">
    <div class="form-container">
      <h2>Purchase Order Form</h2>
      <form id="poForm">
        <div class="section">
          <div class="form-grid">
            <div class="form-field">
              <label>PO Date:</label>
              <input type="date" name="poDate" id="poDate" readonly required>
            </div>
            <div class="form-field">
              <label>PO Number:</label>
              <input type="text" name="poNumber" id="poNumber" required>
            </div>
            <div class="form-field">
              <label>Requisition ID:</label>
              <select name="reqId" id="reqId" required>
                <option value="">Loading Requisitions...</option>
              </select>
            </div>
            <div class="form-field">
              <label>Site:</label>
              <input type="text" name="site" id="site" readonly>
            </div>
            <div class="form-field">
              <label for="poPreparedBy">PO Prepared By:</label>
              <input type="text" id="poPreparedBy" name="poPreparedBy" required>
            </div>
          </div>
        </div>
        <div class="section">
          <h3>Items</h3>
          <table class="line-items-table" id="lineItemsTable">
            <thead>
              <tr>
                <th>Item Name</th><th>Purpose</th><th>Qty</th><th>UOM</th><th>Rate</th><th>GST (%)</th><th>Total Cost</th>
              </tr>
            </thead>
            <tbody id="lineItemsBody">
              <tr><td colspan="7" style="text-align:center;padding:20px;">Select a Requisition ID to see items.</td></tr>
            </tbody>
          </table>
        </div>
        <div class="section">
          <div class="form-grid">
            <div class="form-field">
              <label>Total Cost After GST:</label>
              <input type="number" name="totalCostAfterGST" id="totalCostAfterGST" readonly required>
            </div>
            <div class="form-field">
              <label>Payment Term:</label>
              <input type="text" name="paymentTerm" id="paymentTerm" required>
            </div>
            <div class="form-field">
              <label>Delivery Term:</label>
              <input type="text" name="deliveryTerm" id="deliveryTerm" required>
            </div>
            <div class="form-field">
              <label>Expected Delivery Date:</label>
              <input type="date" name="expectedDeliveryDate" id="expectedDeliveryDate" required>
            </div>
            <div class="form-field">
              <label>Attach PO:</label>
              <input type="file" name="attachPO" required>
            </div>
          </div>
        </div>
        <div class="section" id="vendorDetailsSection">
          <h3>Vendor Details</h3>
          <div class="form-grid">
            <div class="form-field">
              <label>Vendor Registered?</label>
              <input type="text" id="vendorRegistered" name="vendorRegistered" readonly>
            </div>
            <div class="form-field">
              <label>Vendor Company Name:</label>
              <input type="text" id="vendorCompanyName" name="vendorCompanyName" readonly>
            </div>
            <div class="form-field">
              <label>Vendor Contact Person:</label>
              <input type="text" id="vendorContactPerson" name="vendorContactPerson" readonly>
            </div>
            <div class="form-field">
              <label>Vendor Contact Number:</label>
              <input type="text" id="vendorContactNumber" name="vendorContactNumber" readonly>
            </div>
            <div class="form-field">
              <label>Vendor Email:</label>
              <input type="text" id="vendorEmail" name="vendorEmail" readonly>
            </div>
            <div class="form-field">
              <label>Vendor GST Certificate:</label>
              <input type="text" id="vendorGST" name="vendorGST" readonly>
            </div>
            <div class="form-field">
              <label>Vendor PAN Card:</label>
              <input type="text" id="vendorPAN" name="vendorPAN" readonly>
            </div>
            <div class="form-field">
              <label>Cancelled Cheque:</label>
              <input type="text" id="vendorCheque" name="vendorCheque" readonly>
            </div>
          </div>
        </div>
        <button type="submit" id="submitBtn">Submit PO</button>
      </form>
      <div id="msg"></div>
    </div>
  </div>
  <div id="loader"><div style="display:flex;flex-direction:column;align-items:center;"><div class="spinner"></div><span id="loader-text">Loading...</span></div></div>
<script>
  document.addEventListener('DOMContentLoaded', function() {
    document.getElementById('poDate').value = new Date().toISOString().split('T')[0];
    const reqIdSelect = document.getElementById('reqId');
    const submitBtn = document.getElementById('submitBtn');
    submitBtn.disabled = true;

    loadApprovedRequisitions();

    reqIdSelect.addEventListener('change', function() {
      const prID = this.value;
      const tbody = document.getElementById('lineItemsBody');
      tbody.innerHTML = '';
      submitBtn.disabled = true;
      document.getElementById('site').value = '';
      if (!prID) {
        tbody.innerHTML = '<tr><td colspan="7" style="text-align:center;padding:20px;">Select a Requisition ID to see items.</td></tr>';
        return;
      }
      
      showLoader('Fetching Details...');
      google.script.run
        .withSuccessHandler(function(details) {
          if (!details) {
            document.getElementById('msg').textContent = 'Error: Requisition details not found.';
            hideLoader();
            return;
          }
          document.getElementById('paymentTerm').value = details.paymentTerms || '';
          document.getElementById('deliveryTerm').value = details.deliveryTerms || '';
          document.getElementById('expectedDeliveryDate').value = details.expectedDeliveryDate ? new Date(details.expectedDeliveryDate).toISOString().split('T')[0] : '';
          document.getElementById('site').value = details.site || '';

          // Vendor details
          document.getElementById('vendorRegistered').value = details.isVendorRegistered || '';
          document.getElementById('vendorCompanyName').value = (details.vendor && details.vendor['COMPANY NAME']) ? details.vendor['COMPANY NAME'] : '';
          document.getElementById('vendorContactPerson').value = (details.vendor && details.vendor['CONTACT PERSON']) ? details.vendor['CONTACT PERSON'] : '';
          document.getElementById('vendorContactNumber').value = (details.vendor && details.vendor['CONTACT NUMBER']) ? details.vendor['CONTACT NUMBER'] : '';
          document.getElementById('vendorEmail').value = (details.vendor && details.vendor['EMAIL ID']) ? details.vendor['EMAIL ID'] : '';
          document.getElementById('vendorGST').value = (details.vendor && details.vendor['GST Certificate']) ? details.vendor['GST Certificate'] : '';
          document.getElementById('vendorPAN').value = (details.vendor && details.vendor['Pan Card']) ? details.vendor['Pan Card'] : '';
          document.getElementById('vendorCheque').value = (details.vendor && details.vendor['Cancelled Cheque']) ? details.vendor['Cancelled Cheque'] : '';

          if (details.items && details.items.length > 0) {
            details.items.forEach(item => {
              const row = tbody.insertRow();
              row.innerHTML = \`
                <td><input type="text" name="itemName" value="\${item.itemName}" readonly></td>
                <td><input type="text" name="purpose" value="\${item.purpose}" readonly></td>
                <td><input type="number" name="quantity" value="\${item.quantity}" readonly></td>
                <td><input type="text" name="uom" value="\${item.uom}" readonly></td>
                <td><input type="number" name="rate" value="\${item.rate}" readonly></td>
                <td><input type="number" name="gst" value="\${item.gst}" readonly></td>
                <td><input type="number" name="totalCost" value="\${parseFloat(item.totalCost || 0).toFixed(2)}" readonly></td>
              \`;
            });
          } else {
             tbody.innerHTML = '<tr><td colspan="7" style="text-align:center;padding:20px;">No items found for this requisition.</td></tr>';
          }
          updateGrandTotal();
          submitBtn.disabled = false;
          hideLoader();
        })
        .withFailureHandler(function(err) {
          document.getElementById('msg').textContent = 'Error fetching details: ' + err.message;
          hideLoader();
        })
        .getRequisitionDetailsForPO(prID);
    });

    document.getElementById('poForm').addEventListener('submit', handleFormSubmit);
  });

  function loadApprovedRequisitions() {
    const reqIdSelect = document.getElementById('reqId');
    reqIdSelect.innerHTML = '<option value="">Loading Requisitions...</option>';
    google.script.run
      .withSuccessHandler(function(ids) {
        reqIdSelect.innerHTML = '<option value="">Select Requisition ID...</option>';
        if (ids && ids.length > 0) {
          ids.forEach(id => {
            const option = document.createElement('option');
            option.value = id;
            option.textContent = id;
            reqIdSelect.appendChild(option);
          });
        } else {
          reqIdSelect.innerHTML = '<option value="">No Approved Requisitions Found</option>';
        }
      })
      .withFailureHandler(function() {
        reqIdSelect.innerHTML = '<option value="">Error loading requisitions</option>';
      })
      .getApprovedRequisitions();
  }

  function updateGrandTotal() {
    const rows = document.querySelectorAll('#lineItemsBody tr');
    let grandTotal = 0;
    rows.forEach(row => {
      const input = row.querySelector('input[name="totalCost"]');
      if (input) {
        const val = parseFloat(input.value) || 0;
        grandTotal += val;
      }
    });
    document.getElementById('totalCostAfterGST').value = grandTotal.toFixed(2);
  }

  function handleFormSubmit(e) {
    e.preventDefault();
    showLoader('Submitting PO...');
    const form = e.target;
    const submitBtn = document.getElementById('submitBtn');
    submitBtn.disabled = true;

    const formData = {
      reqId: form.reqId.value,
      poDate: form.poDate.value,
      poNumber: form.poNumber.value,
      paymentTerm: form.paymentTerm.value,
      deliveryTerm: form.deliveryTerm.value,
      expectedDeliveryDate: form.expectedDeliveryDate.value,
      poPreparedBy: form.poPreparedBy.value
    };

    const lineItems = [];
    document.querySelectorAll('#lineItemsBody tr').forEach(row => {
      lineItems.push({
        itemName: row.querySelector('[name="itemName"]').value,
        purpose: row.querySelector('[name="purpose"]').value,
        quantity: row.querySelector('[name="quantity"]').value,
        uom: row.querySelector('[name="uom"]').value,
        rate: row.querySelector('[name="rate"]').value,
        gst: row.querySelector('[name="gst"]').value,
        totalCost: row.querySelector('[name="totalCost"]').value
      });
    });

    const fileInput = form.querySelector('input[name="attachPO"]');
    const filePromise = fileInput && fileInput.files.length > 0
      ? toBase64(fileInput.files[0]).then(base64 => new Promise(resolve => {
          google.script.run
            .withSuccessHandler(url => resolve(url))
            .withFailureHandler(() => resolve(''))
            .uploadFile({ name: fileInput.files[0].name, mimeType: fileInput.files[0].type, base64: base64 }, 'PO_Attachments');
        }))
      : Promise.resolve('');

    filePromise.then(url => {
      formData.attachPO = url;
      google.script.run
        .withSuccessHandler(resp => {
          hideLoader();
          const msgDiv = document.getElementById('msg');
          if (resp && resp.success) {
            msgDiv.textContent = 'PO submitted successfully! PO ID: ' + resp.poID;
            msgDiv.className = 'success';
            form.reset();
            document.getElementById('lineItemsBody').innerHTML = '<tr><td colspan="7" style="text-align:center;padding:20px;">Select a Requisition ID to see items.</td></tr>';
            document.getElementById('poDate').value = new Date().toISOString().split('T')[0];
            document.getElementById('totalCostAfterGST').value = '';
            loadApprovedRequisitions(); // Refresh the requisition list
          } else {
            msgDiv.textContent = 'Submission failed. Please try again.';
            msgDiv.className = 'error';
            submitBtn.disabled = false;
          }
        })
        .withFailureHandler(err => {
          hideLoader();
          document.getElementById('msg').textContent = 'Submission failed: ' + err.message;
          document.getElementById('msg').className = 'error';
          submitBtn.disabled = false;
        })
        .submitPO(formData, lineItems);
    });
  }

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
</body>
</html>
  `;
}