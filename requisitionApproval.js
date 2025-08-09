function requisitionApprovalPage(data) {
    // data is now a JSON string
    const pendingRequisitions = JSON.parse(data.pendingRequisitions || '[]');

    let listHtml = '';
    if (pendingRequisitions.length === 0) {
        listHtml = '<p class="no-pending">No pending requisitions for approval.</p>';
    } else {
        listHtml = pendingRequisitions.map(req => `
            <div class="requisition-card" id="card-${req.id}">
                <div class="card-left">
                    <div class="card-id-row">
                        <span class="pr-id">${req.id}</span>
                        <span class="requested-by">by ${req.requestedBy}</span>
                    </div>
                    <div class="card-details-row">
                        <span class="site">Site: ${req.site}</span>
                        <span class="submission-date">Submitted: ${req.date ? new Date(req.date).toLocaleDateString() : 'N/A'}</span>
                        <span class="item-count">${req.items ? req.items.length : 0} items</span>
                    </div>
                </div>
                <div class="card-actions">
                    <button class="view-details-btn" onclick="window.showDetails('${req.id}')">View Details</button>
                </div>
            </div>
        `).join('');
    }

    return `
    <style>
        /* New Styles for the list view */
        .requisition-list {
            display: grid;
            gap: 15px;
        }
        .requisition-card {
            background-color: var(--card-background);
            border-radius: 8px;
            box-shadow: var(--shadow);
            padding: 15px 20px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-left: 5px solid #007bff; /* A new color */
            transition: all 0.2s ease;
        }
        .requisition-card:hover {
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
            transform: translateY(-2px);
        }
        .card-left {
            display: flex;
            flex-direction: column;
        }
        .card-id-row {
            display: flex;
            align-items: center;
            gap: 10px;
            font-size: 1.1rem;
        }
        .pr-id {
            font-weight: 600;
            color: #007bff;
        }
        .requested-by {
            color: var(--text-color);
            font-weight: 500;
        }
        .card-details-row {
            display: flex;
            align-items: center;
            gap: 20px;
            font-size: 0.9rem;
            color: var(--subtle-text-color);
            margin-top: 5px;
        }
        .submission-date, .site, .item-count {
            font-style: italic;
        }
        .view-details-btn {
            background-color: #e6f0ff;
            color: #007bff;
            border: none;
            border-radius: 5px;
            padding: 8px 16px;
            cursor: pointer;
            font-weight: 500;
            transition: all 0.2s;
        }
        .view-details-btn:hover {
            background-color: #007bff;
            color: white;
        }
        
        /* Modal Styles */
        .modal-overlay {
            position: fixed; top: 0; left: 0; width: 100%; height: 100%;
            background: rgba(0,0,0,0.6); display: none; align-items: center; justify-content: center; z-index: 1000;
        }
        .modal-content {
            background: white; border-radius: 8px; width: 90%; max-width: 800px;
            max-height: 90vh; overflow-y: auto; box-shadow: 0 5px 15px rgba(0,0,0,0.3);
            padding: 30px;
        }
        .modal-header {
            padding-bottom: 15px;
            border-bottom: 1px solid var(--border-color);
            display: flex; justify-content: space-between; align-items: center;
        }
        .modal-header h2 { margin: 0; font-size: 1.4rem; color: #007bff; }
        .modal-close { font-size: 1.8rem; cursor: pointer; color: var(--subtle-text-color); border: none; background: none; }
        .modal-body { padding-top: 25px; }
        .modal-grid { 
            display: grid; 
            grid-template-columns: 1fr 1fr; 
            gap: 25px; 
            margin-bottom: 20px;
            background-color: #f8f9fa;
            padding: 20px;
            border-radius: 8px;
        }
        .modal-section h4 { 
            margin: 0 0 10px 0; 
            color: #333; 
            border-bottom: 1px solid #eee; 
            padding-bottom: 5px; 
            font-weight: 600;
        }
        .modal-section p { 
            margin: 5px 0; 
            font-size: 0.95rem; 
            display: flex;
            gap: 10px;
        }
        .modal-section p strong {
            color: var(--subtle-text-color);
            font-weight: 500;
            min-width: 150px;
        }
        .items-table { width: 100%; border-collapse: collapse; margin-top: 20px; }
        .items-table th, .items-table td { padding: 12px; text-align: left; border-bottom: 1px solid var(--border-color); }
        .items-table th { background-color: #e9ecef; font-weight: 600; color: #555; }
        .total-amount { 
            text-align: right; 
            font-size: 1.2rem; 
            font-weight: bold; 
            margin-top: 15px;
            padding: 10px 0;
            border-top: 2px solid #333;
        }
        .modal-footer { 
            padding-top: 20px; 
            text-align: right; 
            border-top: 1px solid var(--border-color);
        }
        #modal-remarks { 
            width: 100%; 
            padding: 10px; 
            border: 1px solid var(--border-color); 
            border-radius: 4px; 
            min-height: 80px; 
            margin-bottom: 15px;
        }
        .no-pending { text-align: center; padding: 40px; font-size: 1.1rem; color: var(--subtle-text-color); cursor: default; }
        .btn {
            border: none;
            border-radius: 5px;
            color: white;
            cursor: pointer;
            font-weight: 500;
            font-size: 1rem;
            padding: 10px 20px;
            transition: background-color 0.2s;
            margin-left: 10px;
        }
        .btn-approve { background-color: #28a745; }
        .btn-reject { background-color: #dc3545; }

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
        #msg {
            margin-top: 20px;
            font-weight: bold;
            padding: 10px;
            border-radius: 5px;
            text-align: center;
            display: none;
        }
        .success { color: #22c55e; background-color: #dcfce7; }
        .error { color: #ef4444; background-color: #fee2e2; }

    </style>

    <div class="container">
        <h1>Pending Approvals</h1>
        <div id="requisition-list" class="requisition-list">${listHtml}</div>
        <div id="msg"></div>
    </div>

    <div class="modal-overlay" id="details-modal">
        <div class="modal-content">
            <div class="modal-header">
                <h2 id="modal-title"></h2>
                <button class="modal-close" onclick="window.hideModal()">&times;</button>
            </div>
            <div class="modal-body">
                <div class="modal-grid">
                    <div class="modal-section">
                        <h4>Requisition Information</h4>
                        <p><strong>ID:</strong> <span id="modal-req-id"></span></p>
                        <p><strong>Employee:</strong> <span id="modal-req-employee"></span></p>
                        <p><strong>Site:</strong> <span id="modal-req-site"></span></p>
                        <p><strong>Date:</strong> <span id="modal-req-date"></span></p>
                        <p><strong>Expected Delivery:</strong> <span id="modal-req-delivery-date"></span></p>
                        <p><strong>Status:</strong> <span id="modal-req-status" style="color: #f0ad4e; font-weight: bold;"></span></p>
                        <p><strong>Purchase Category:</strong> <span id="modal-req-category"></span></p>
                    </div>
                    <div class="modal-section">
                        <h4>Vendor Information</h4>
                        <p><strong>Company Name:</strong> <span id="modal-vendor-company"></span></p>
                        <p><strong>Contact Person:</strong> <span id="modal-vendor-contact-person"></span></p>
                        <p><strong>Email:</strong> <span id="modal-vendor-email"></span></p>
                        <p><strong>Phone:</strong> <span id="modal-vendor-phone"></span></p>
                    </div>
                </div>
                <div class="modal-section">
                    <h4>Items</h4>
                    <table class="items-table">
                        <thead><tr><th>Item</th><th>Quantity</th><th>Unit Price</th><th>Total</th></tr></thead>
                        <tbody id="modal-items-table"></tbody>
                    </table>
                    <div class="total-amount" id="modal-total-amount"></div>
                </div>
                <div class="modal-section">
                    <h4>Approver Remarks</h4>
                    <textarea id="modal-remarks" placeholder="Remarks are required for rejection..."></textarea>
                </div>
            </div>
            <div class="modal-footer">
                <button class="btn btn-reject" id="modal-reject-btn" onclick="window.handleApproval(prID, 'Rejected')">Decline</button>
                <button class="btn btn-approve" id="modal-approve-btn" onclick="window.handleApproval(prID, 'Approved')">Approve</button>
            </div>
        </div>
    </div>

    <div id="loader"><div style="display:flex;flex-direction:column;align-items:center;"><div class="spinner"></div><span id="loader-text">Loading...</span></div></div>
    <script>
        const requisitions = ${JSON.stringify(pendingRequisitions)};
        const modal = document.getElementById('details-modal');
        const msgDiv = document.getElementById('msg');
        const loader = document.getElementById('loader');

        window.showDetails = function(prID) {
            const req = requisitions.find(r => r.id === prID);
            if (!req) return;

            document.getElementById('modal-title').innerText = 'Requisition Details: ' + prID;
            
            // Requisition Info
            document.getElementById('modal-req-id').innerText = req.id || 'N/A';
            document.getElementById('modal-req-employee').innerText = req.requestedBy || 'N/A';
            document.getElementById('modal-req-site').innerText = req.site || 'N/A';
            document.getElementById('modal-req-date').innerText = req.date ? new Date(req.date).toLocaleDateString() : 'N/A';
            document.getElementById('modal-req-delivery-date').innerText = req.expectedDeliveryDate ? new Date(req.expectedDeliveryDate).toLocaleDateString() : 'N/A';
            document.getElementById('modal-req-status').innerText = req.status || 'N/A';
            document.getElementById('modal-req-category').innerText = req.purchaseCategory || 'N/A';

            // Vendor Info
            document.getElementById('modal-vendor-company').innerText = (req.vendor && req.vendor.companyName) ? req.vendor.companyName : 'N/A';
            document.getElementById('modal-vendor-contact-person').innerText = (req.vendor && req.vendor.contactPerson) ? req.vendor.contactPerson : 'N/A';
            document.getElementById('modal-vendor-email').innerText = (req.vendor && req.vendor.email) ? req.vendor.email : 'N/A';
            document.getElementById('modal-vendor-phone').innerText = (req.vendor && req.vendor.contactNumber) ? req.vendor.contactNumber : 'N/A';


            // Items Table
            const itemsTbody = document.getElementById('modal-items-table');
            itemsTbody.innerHTML = (req.items && req.items.length > 0)
              ? req.items.map(item => \`
                <tr>
                    <td>\${item.itemName || 'N/A'}</td>
                    <td>\${item.quantity || 'N/A'} \${item.uom || 'N/A'}</td>
                    <td>₹\${parseFloat(item.rate).toFixed(2)}</td>
                    <td>₹\${parseFloat(item.totalValue).toFixed(2)}</td>
                </tr>
              \`).join('')
              : '<tr><td colspan="4" style="text-align:center;">No items found for this requisition.</td></tr>';

            document.getElementById('modal-total-amount').innerText = 'Total Amount: ₹' + parseFloat(req.totalValue).toFixed(2);
            
            // Set up modal buttons
            document.getElementById('modal-approve-btn').onclick = () => window.handleApproval(prID, 'Approved');
            document.getElementById('modal-reject-btn').onclick = () => window.handleApproval(prID, 'Rejected');

            modal.style.display = 'flex';
        }

        window.hideModal = function() {
            modal.style.display = 'none';
            document.getElementById('modal-remarks').value = '';
        }

        window.showLoader = function(text) {
            document.getElementById('loader-text').textContent = text;
            loader.style.display = 'flex';
        }

        window.hideLoader = function() {
            loader.style.display = 'none';
        }


        window.handleApproval = function(prID, action, fromCard = false) {
            const remarks = fromCard ? '' : document.getElementById('modal-remarks').value;
            
            if (action === 'Rejected' && !remarks && !fromCard) {
                alert('Remarks are required for rejection.');
                return;
            }

            const card = document.getElementById('card-' + prID);
            showLoader(\`Processing PR \${prID}...\`);
            
            if (!fromCard) window.hideModal();

            google.script.run
                .withSuccessHandler(response => {
                    hideLoader();
                    msgDiv.className = 'success';
                    msgDiv.innerText = \`Requisition \${prID} has been successfully \${action}.\`;
                    msgDiv.style.display = 'block';

                    card.style.transition = 'opacity 0.5s ease';
                    card.style.opacity = '0';
                    setTimeout(() => {
                        card.remove();
                        if (document.querySelectorAll('.requisition-card').length === 0) {
                            document.getElementById('requisition-list').innerHTML = '<div class="card no-pending"><p>No requisitions are currently pending approval.</p></div>';
                        }
                    }, 500);
                })
                .withFailureHandler(error => {
                    hideLoader();
                    msgDiv.className = 'error';
                    msgDiv.innerText = 'An error occurred: ' + error.message;
                    msgDiv.style.display = 'block';
                    card.classList.removef('processing');
                })
                .processApproval(prID, action, remarks);
        }
        
        // Close modal on overlay click
        window.onclick = function(event) {
            if (event.target == modal) {
                window.hideModal();
            }
        }
    </script>
    `;
}