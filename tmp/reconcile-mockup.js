const fs = require('fs');
const path = require('path');

const file = path.join(__dirname, '..', 'form_submit_tool.html');
let html = fs.readFileSync(file, 'utf8');

html = html.replace(/state: "awaiting_gateway"/g, 'state: "awaiting"');
html = html.replace(/state: "not_required"/g, 'state: "not_configured"');

// Multiline legacy blob (remove whole object, not just the key line)
html = html.replace(/^                filesDetail: \{[\s\S]*?^                \},\r?\n/gm, '');

// Top-level record scalar fields only (16-space indent)
const deadTop = /^                (pdf|pdfBool|files|approval|approvalTone|payment|paymentRequired|paymentComplete|reconcileNeeded):[^\n]*\r?\n/gm;
html = html.replace(deadTop, '');

html = html.replace(/^                    paymentComplete: (true|false),\r?\n/gm, '');
html = html.replace(/^                    reconcileNeeded: (true|false),\r?\n/gm, '');

html = html.replace(
    /payment: \{ required: true, bypassed: true, bypassField: "payment_ty", bypassValue: "K", bypassPage: "pay\.asp" \}/,
    'payment: { required: true, bypassed: true, bypassField: "payment_ty", bypassValue: "K", bypassPage: "pay.asp", state: "not_needed" }'
);

html = html.replace(
    /payment: \{ required: false, method: "none", state: "not_configured" \}/g,
    'payment: { required: false, configured: false, state: "not_configured" }'
);

html = html.replace(
    /\{ name: "ce_block_c\.pdf", size: "[^"]*", state: "failed"/,
    '{ name: "ce_block_c.pdf", size: "198 KB", state: "failed"'
);

fs.writeFileSync(file, html);
console.log('mockup data reconciled');
