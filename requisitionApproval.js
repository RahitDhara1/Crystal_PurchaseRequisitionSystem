function requisitionApprovalPage(data) {
    // data is now a JSON string
    const pendingRequisitions = JSON.parse(data.pendingRequisitions || '[]');

    let listHtml = '';
    if (pendingRequisitions.length === 0) {
        listHtml = '<p class="no-pending">No pending requisitions for approval.</p>';
    } else {
        listHtml = pendingRequisitions.map(req => `
            <div class="requisition-card">
                <div class="card-info">
                    <span class="pr-id">#${req.id}</span>
                    <span class="requested-by">by ${req.requestedBy}</span>
                    <span class="site">for ${req.site}</span>
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
            border-left: 5px solid var(--primary-color);
            transition: all 0.2s ease;
        }
        .requisition-card:hover {
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
            transform: translateY(-2px);
        }
        .card-info {
            display: flex;
            align-items: center;
            gap: 15px;
            font-size: 1rem;
        }
        .pr-id {
            font-weight: 600;
            color: var(--primary-color);
        }
        .requested-by {
            color: var(--text-color);
        }
        .site {
            color: var(--subtle-text-color);
            font-style: italic;
        }
        .view-details-btn {
            background-color: #f0f4ff;
            color: var(--primary-color);
            border: 1px solid #e0e7ff;
            border-radius: 5px;
            padding: 8px 16px;
            cursor: pointer;
            font-weight: 500;
            transition: all 0.2s;
        }
        .view-details-btn:hover {
            background-color: var(--primary-color);
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
        }
        .modal-header {
            padding: 15px 25px; border-bottom: 1px solid var(--border-color);
            display: flex; justify-content: space-between; align-items: center;
        }
        .modal-header h2 { margin: 0; font-size: 1.4rem; }
        .modal-close { font-size: 1.8rem; cursor: pointer; color: var(--subtle-text-color); border: none; background: none; }
        .modal-body { padding: 25px; }
        .modal-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 25px; margin-bottom: 20px; }
        .modal-section h4 { margin: 0 0 10px 0; color: var(--primary-color); border-bottom: 1px solid #eee; padding-bottom: 5px; }
        .modal-section p { margin: 5px 0; font-size: 0.95rem; }
        .items-table { width: 100%; border-collapse: collapse; margin-top: 10px; }
        .items-table th, .items-table td { padding: 10px; text-align: left; border-bottom: 1px solid var(--border-color); }
        .items-table th { background-color: #f8f9fa; font-weight: 500; }
        .items-table td:nth-child(3), .items-table td:nth-child(4) { text-align: right; }
        .total-amount { text-align: right; font-size: 1.2rem; font-weight: bold; margin-top: 15px; }
        .modal-footer { padding: 15px 25px; text-align: right; background-color: #fafafa; border-top: 1px solid var(--border-color); }
        #modal-remarks { width: 100%; padding: 8px; border: 1px solid var(--border-color); border-radius: 4px; min-height: 50px; margin-bottom: 15px; }
        .no-pending { text-align: center; padding: 40px; font-size: 1.1rem; color: var(--subtle-text-color); cursor: default; }
    </style>

    <div class="container">
        <h1>Pending Approvals</h1>
        <div id="requisition-list" class="requisition-list">${listHtml}</div>
    </div>

    <div class="modal-overlay" id="details-modal">
        <div class="modal-content">
            <div class="modal-header">
                <h2 id="modal-title">Requisition Details</h2>
                <button class="modal-close" onclick="window.hideModal()">&times;</button>
            </div>
            <div class="modal-body">
                <div class="modal-grid">
                    <div class="modal-section" id="modal-req-info"></div>
                    <div class="modal-section" id="modal-vendor-info"></div>
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
                <button class="btn btn-reject" id="modal-reject-btn">Decline</button>
                <button class="btn btn-approve" id="modal-approve-btn">Approve</button>
            </div>
        </div>
    </div>

    <script>
        const requisitions = ${JSON.stringify(pendingRequisitions)};
        const modal = document.getElementById('details-modal');

        window.showDetails = function(prID) {
            const req = requisitions.find(r => r.id === prID);
            if (!req) return;

            document.getElementById('modal-title').innerText = 'Requisition Details: ' + prID;
            
            // Requisition Info
            document.getElementById('modal-req-info').innerHTML = \`
                <h4>Requisition Information</h4>
                <p><strong>ID:</strong> \${req.id}</p>
                <p><strong>Employee:</strong> \${req.requestedBy}</p>
                <p><strong>Site:</strong> \${req.site || 'N/A'}</p>
                <p><strong>Date:</strong> \${req.date ? new Date(req.date).toLocaleDateString() : 'N/A'}</p>
                <p><strong>Expected Delivery:</strong> \${req.expectedDeliveryDate ? new Date(req.expectedDeliveryDate).toLocaleDateString() : 'N/A'}</p>
                <p><strong>Status:</strong> <span style="color: #f0ad4e; font-weight: bold;">\${req.status}</span></p>
            \`;

            // Vendor Info
            document.getElementById('modal-vendor-info').innerHTML = \`
                <h4>Vendor Information</h4>
                <p><strong>Company Name:</strong> \${req.vendor.name || req.vendor.companyName || 'N/A'}</p>
                <p><strong>Contact Person:</strong> \${req.vendor.contactPerson || 'N/A'}</p>
                <p><strong>Email:</strong> \${req.vendor.email || 'N/A'}</p>
                <p><strong>Phone:</strong> \${req.vendor.contactNumber || 'N/A'}</p>
            \`;

            // Items Table
            const itemsTbody = document.getElementById('modal-items-table');
            itemsTbody.innerHTML = (req.items && req.items.length > 0)
              ? req.items.map(item => \`
                <tr>
                    <td>\${item.itemName}</td>
                    <td>\${item.quantity} \${item.uom}</td>
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

        window.handleApproval = function(prID, action, fromCard = false) {
            const remarks = fromCard ? '' : document.getElementById('modal-remarks').value;
            
            if (action === 'Rejected' && !remarks && !fromCard) {
                alert('Remarks are required for rejection.');
                return;
            }

            const card = document.getElementById('card-' + prID);
            card.classList.add('processing');
            if (!fromCard) window.hideModal();

            google.script.run
                .withSuccessHandler(response => {
                    card.style.transition = 'opacity 0.5s ease';
                    card.style.opacity = '0';
                    setTimeout(() => {
                        card.remove();
                        if (document.querySelectorAll('.card:not(.no-pending)').length === 0) {
                            document.getElementById('requisition-list').innerHTML = '<div class="card no-pending"><p>No requisitions are currently pending approval.</p></div>';
                        }
                    }, 500);
                })
                .withFailureHandler(error => {
                    alert('An error occurred: ' + error.message);
                    card.classList.remove('processing');
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