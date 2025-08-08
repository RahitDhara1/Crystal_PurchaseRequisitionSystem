function poApprovalPage(data) {
    // data is now a JSON string
    const pendingPOs = JSON.parse(data.pendingPOs || '[]');

    let listHtml = '';
    if (pendingPOs.length === 0) {
        listHtml = '<p>No pending purchase orders for approval.</p>';
    } else {
        listHtml = pendingPOs.map(po => `
            <div class="po-card">
                <div class="card-header">
                    <h3>PO No: ${po.poId} (PR: ${po.prId})</h3>
                </div>
                <div class="card-body">
                    <p><strong>Requested By:</strong> ${po.requestedBy}</p>
                    <p><strong>Site:</strong> ${po.site}</p>
                    <p><strong>Date:</strong> ${new Date(po.date).toLocaleDateString()}</p>
                    <p><strong>Total Value:</strong> ${po.totalValue}</p>
                    <p><strong>Vendor:</strong> ${po.vendor ? po.vendor.companyName : 'N/A'}</p>
                </div>
                <div class="card-actions">
                    <button onclick="showDetails('${po.poId}')">View Details</button>
                </div>
            </div>
        `).join('');
    }

    return `
    <style>
        /* Page-specific styles for approval page */
        .container { max-width: 1100px; margin: 0 auto; padding: 0; }
        h1 { color: var(--text-color); margin-bottom: 20px; }
        .card {
            background-color: var(--card-background); border-radius: 8px; box-shadow: var(--shadow);
            margin-bottom: 15px; display: flex; align-items: center; padding: 20px;
            transition: all 0.3s ease; cursor: pointer;
        }
        .card:hover { transform: translateY(-3px); box-shadow: 0 6px 16px rgba(0,0,0,0.1); }
        .card.processing { opacity: 0.5; pointer-events: none; }
        .card-main { flex-grow: 1; }
        .card-main h4 { margin: 0 0 8px 0; color: var(--primary-color); }
        .card-main p { margin: 4px 0; font-size: 0.9rem; color: var(--subtle-text-color); }
        .card-actions .btn { padding: 8px 16px; margin-left: 10px; }
        .btn {
            border: none; border-radius: 5px; color: white; cursor: pointer;
            font-weight: 500; font-size: 0.9rem; transition: background-color 0.2s;
        }
        .btn-approve { background-color: var(--green); }
        .btn-reject { background-color: var(--red); }
        .no-pending { text-align: center; padding: 40px; font-size: 1.1rem; color: var(--subtle-text-color); cursor: default; }
        
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
    </style>

    <div class="container">
        <h1>Pending PO Approvals</h1>
        <div id="po-list">${listHtml}</div>
    </div>

    <!-- Modal Structure -->
    <div class="modal-overlay" id="details-modal">
        <div class="modal-content">
            <div class="modal-header">
                <h2 id="modal-title">PO Details</h2>
                <button class="modal-close" onclick="hideModal()">&times;</button>
            </div>
            <div class="modal-body">
                <div class="modal-grid">
                    <div class="modal-section" id="modal-po-info"></div>
                    <div class="modal-section" id="modal-vendor-info"></div>
                </div>
                <div class="modal-section">
                    <h4>Items</h4>
                    <table class="items-table">
                        <thead><tr><th>Item</th><th>Quantity</th><th>UOM</th><th>Unit Price</th><th>Total</th></tr></thead>
                        <tbody id="modal-items-table"></tbody>
                    </table>
                    <div class="total-amount" id="modal-total-amount"></div>
                </div>
                 <div class="modal-section">
                    <h4>Attachments</h4>
                    <p><strong>Attached PO:</strong> <a id="modal-po-attachment" href="#" target="_blank">View Attachment</a></p>
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
        const pos = ${JSON.stringify(pendingPOs)};
        const modal = document.getElementById('details-modal');

        function showModal(poID) {
            const po = pos.find(p => p.id === poID);
            if (!po) return;

            document.getElementById('modal-title').innerText = 'PO Details: ' + poID;
            
            // PO Info
            document.getElementById('modal-po-info').innerHTML = \`
                <h4>PO Information</h4>
                <p><strong>PO ID:</strong> \${po.id}</p>
                <p><strong>Requisition ID:</strong> \${po.reqId}</p>
                <p><strong>PO Date:</strong> \${po.date ? new Date(po.date).toLocaleDateString() : 'N/A'}</p>
                <p><strong>Site:</strong> \${po.site || 'N/A'}</p>
                <p><strong>Status:</strong> <span style="color: #f0ad4e; font-weight: bold;">\${po.status}</span></p>
            \`;

            // Vendor Info
            document.getElementById('modal-vendor-info').innerHTML = \`
                <h4>Vendor Information</h4>
                <p><strong>Company Name:</strong> \${po.vendor.name || 'N/A'}</p>
                <p><strong>Contact Person:</strong> \${po.vendor.contactPerson || 'N/A'}</p>
                <p><strong>Email:</strong> \${po.vendor.email || 'N/A'}</p>
                <p><strong>Phone:</strong> \${po.vendor.contactNumber || 'N/A'}</p>
            \`;

            // Items Table
            const itemsTbody = document.getElementById('modal-items-table');
            itemsTbody.innerHTML = (po.items && po.items.length > 0)
              ? po.items.map(item => \`
                <tr>
                    <td>\${item.itemName}</td>
                    <td>\${item.quantity}</td>
                    <td>\${item.uom}</td>
                    <td>₹\${parseFloat(item.rate).toFixed(2)}</td>
                    <td>₹\${parseFloat(item.totalCost).toFixed(2)}</td>
                </tr>
              \`).join('')
              : '<tr><td colspan="5" style="text-align:center;">No items found for this PO.</td></tr>';

            document.getElementById('modal-total-amount').innerText = 'Total Amount: ₹' + parseFloat(po.totalValue).toFixed(2);
            
            // Attachment
            const poAttachmentLink = document.getElementById('modal-po-attachment');
            if (po.attachPO) {
                poAttachmentLink.href = po.attachPO;
                poAttachmentLink.style.display = 'inline';
            } else {
                poAttachmentLink.style.display = 'none';
            }

            // Set up modal buttons
            document.getElementById('modal-approve-btn').onclick = () => handleApproval(poID, 'Approved');
            document.getElementById('modal-reject-btn').onclick = () => handleApproval(poID, 'Rejected');

            modal.style.display = 'flex';
        }

        function hideModal() {
            modal.style.display = 'none';
            document.getElementById('modal-remarks').value = '';
        }

        function handleApproval(poID, action, fromCard = false) {
            const remarks = fromCard ? '' : document.getElementById('modal-remarks').value;
            
            if (action === 'Rejected' && !remarks && !fromCard) {
                alert('Remarks are required for rejection.');
                return;
            }

            const card = document.getElementById('card-' + poID);
            card.classList.add('processing');
            if (!fromCard) hideModal();

            google.script.run
                .withSuccessHandler(response => {
                    card.style.transition = 'opacity 0.5s ease';
                    card.style.opacity = '0';
                    setTimeout(() => {
                        card.remove();
                        if (document.querySelectorAll('.card:not(.no-pending)').length === 0) {
                            document.getElementById('po-list').innerHTML = '<div class="card no-pending"><p>No purchase orders are currently pending approval.</p></div>';
                        }
                    }, 500);
                })
                .withFailureHandler(error => {
                    alert('An error occurred: ' + error.message);
                    card.classList.remove('processing');
                })
                .processPOApproval(poID, action, remarks);
        }
        
        // Close modal on overlay click
        window.onclick = function(event) {
            if (event.target == modal) {
                hideModal();
            }
        }
    </script>
    `;
}
