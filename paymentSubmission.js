function paymentSubmissionPage() {
  return `
    <div class="container">
      <h1>Payment Submission</h1>
      <form id="paymentForm">
        <div class="form-group">
          <label for="poNumber">PO Number:</label>
          <input type="text" id="poNumber" name="poNumber" required>
        </div>
        <div class="form-group">
          <label for="amountPaid">Amount Paid:</label>
          <input type="number" id="amountPaid" name="amountPaid" required>
        </div>
        <div class="form-group">
          <label for="paymentDate">Payment Date:</label>
          <input type="date" id="paymentDate" name="paymentDate" required>
        </div>
        <div class="form-group">
          <label for="paymentMode">Payment Mode:</label>
          <select id="paymentMode" name="paymentMode" required>
            <option value="NEFT">NEFT</option>
            <option value="RTGS">RTGS</option>
            <option value="IMPS">IMPS</option>
            <option value="UPI">UPI</option>
            <option value="Check">Check</option>
            <option value="Cash">Cash</option>
          </select>
        </div>
        <div class="form-group">
          <label for="utrNumber">UTR Number:</label>
          <input type="text" id="utrNumber" name="utrNumber">
        </div>
        <div class="form-group">
          <label for="remarks">Remarks:</label>
          <textarea id="remarks" name="remarks"></textarea>
        </div>
        <button type="button" onclick="submitForm()">Submit Payment</button>
      </form>
    </div>

    <script>
      function submitForm() {
        const form = document.getElementById('paymentForm');
        const formData = {
          poNumber: form.poNumber.value,
          amountPaid: form.amountPaid.value,
          paymentDate: form.paymentDate.value,
          paymentMode: form.paymentMode.value,
          utrNumber: form.utrNumber.value,
          remarks: form.remarks.value
        };

        google.script.run
          .withSuccessHandler(function(response) {
            alert(response.message);
            form.reset();
          })
          .withFailureHandler(function(err) {
            alert('An error occurred: ' + err.message);
          })
          .submitPayment(formData);
      }
    </script>
  `;
}