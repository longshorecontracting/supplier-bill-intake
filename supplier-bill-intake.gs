// =================================================================================================
// SUPPLIER BILL INTAKE PIPELINE — SANITIZED TEMPLATE
// Purpose:
//   1. Watch Gmail label for supplier invoice PDFs
//   2. Parse PDFs with Google Document AI
//   3. Write one row per invoice line item into a Google Sheets Master tab
//   4. Move processed/failed emails to the correct Gmail labels
//   5. Optionally forward PDFs to accounting inbox/QBO
// =================================================================================================

var SETTINGS = {
  // Replace with your Google Document AI processor endpoint.
  // Format usually looks like:
  // https://us-documentai.googleapis.com/v1/projects/YOUR_PROJECT_ID/locations/us/processors/YOUR_PROCESSOR_ID:process
  PREDICTION_ENDPOINT: 'REPLACE_WITH_DOCUMENT_AI_PROCESSOR_ENDPOINT',

  // Gmail labels. Create these in Gmail first, or let the script create processed/error labels.
  GMAIL_LABEL_IN:      'Suppliers/VENDOR_NAME/To-Process',
  GMAIL_LABEL_TO_FILE: 'Suppliers/VENDOR_NAME/To-File',
  GMAIL_LABEL_ERROR:   'Suppliers/VENDOR_NAME/Errors',

  // Target Google Sheet.
  SPREADSHEET_ID:      'REPLACE_WITH_GOOGLE_SHEET_ID',
  SHEET_NAME:          'Master',

  // Optional accounting inbox. Leave blank to disable forwarding.
  ACCOUNTING_FORWARD_EMAIL: '',

  // Safety switch. Keep true while testing.
  PAUSE_ACCOUNTING_EMAILS: true
};

function processSupplierInvoices() {
  var labelIn = GmailApp.getUserLabelByName(SETTINGS.GMAIL_LABEL_IN);
  if (!labelIn) throw new Error('Missing Gmail label: ' + SETTINGS.GMAIL_LABEL_IN);

  var threads = labelIn.getThreads(0, 50);
  if (threads.length === 0) {
    console.log('No invoice threads to process.');
    return;
  }

  var ss = SpreadsheetApp.openById(SETTINGS.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SETTINGS.SHEET_NAME) || ss.insertSheet(SETTINGS.SHEET_NAME);
  ensureMasterHeader_(sheet);

  var existing = buildInvoiceKeySet_(sheet);
  var labelToFile = getOrCreateLabel_(SETTINGS.GMAIL_LABEL_TO_FILE);
  var labelError = getOrCreateLabel_(SETTINGS.GMAIL_LABEL_ERROR);

  var successfulThreads = [];
  var processedPDFs = [];
  var processedCount = 0;

  var skipKeywords = [
    'STATEMENT', 'PACKING SLIP', 'PACKING SLIPS', 'DELIVERY RECEIPT',
    'RECEIPT', 'SLIP', 'CREDIT MEMO', 'CREDIT', 'MEMO'
  ];

  for (var t = 0; t < threads.length; t++) {
    var thread = threads[t];
    var threadSuccess = false;
    var messages = thread.getMessages();

    for (var m = 0; m < messages.length; m++) {
      var msg = messages[m];
      var pdfs = msg.getAttachments({ mimeType: 'application/pdf' });
      if (!pdfs || pdfs.length === 0) continue;

      for (var p = 0; p < pdfs.length; p++) {
        var pdf = pdfs[p];
        var pdfNameUpper = pdf.getName().toUpperCase();

        var shouldSkip = skipKeywords.some(function(keyword) {
          return pdfNameUpper.indexOf(keyword) !== -1;
        });
        if (shouldSkip) {
          console.log('Skipping non-invoice PDF: ' + pdf.getName());
          continue;
        }

        try {
          var parsed = parsePdfWithDocumentAi_(pdf);
          var invoiceKey = parsed.vendor + '|||' + parsed.invoiceNumber;

          if (existing.has(invoiceKey)) {
            console.log('Already imported, skipping invoice: ' + parsed.invoiceNumber);
            continue;
          }

          var rows = buildMasterRows_(parsed, msg.getId());
          if (rows.length > 0) {
            sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
            existing.add(invoiceKey);
            processedCount++;
            processedPDFs.push({ attachment: pdf, invoiceNumber: parsed.invoiceNumber || pdf.getName() });
            threadSuccess = true;
            console.log('Imported invoice ' + parsed.invoiceNumber + ' with ' + rows.length + ' rows.');
          }
        } catch (err) {
          console.error('Failed PDF: ' + pdf.getName() + ' — ' + err.message);
          labelError.addToThread(thread);
        }

        Utilities.sleep(1000);
      }
    }

    if (threadSuccess) {
      successfulThreads.push(thread);
    }
  }

  if (processedPDFs.length > 0) {
    maybeForwardPdfsToAccounting_(processedPDFs);

    successfulThreads.forEach(function(thread) {
      thread.removeLabel(labelIn);
      labelToFile.addToThread(thread);
    });
  }

  console.log('Done. Processed invoices: ' + processedCount);
}

function parsePdfWithDocumentAi_(pdf) {
  var pdfBytes = pdf.getBytes();
  if (pdfBytes.length > 20 * 1024 * 1024) {
    throw new Error('PDF too large: ' + Math.round(pdfBytes.length / 1024 / 1024) + 'MB');
  }

  var response = UrlFetchApp.fetch(SETTINGS.PREDICTION_ENDPOINT, {
    method: 'post',
    muteHttpExceptions: true,
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
    },
    payload: JSON.stringify({
      rawDocument: {
        content: Utilities.base64Encode(pdfBytes),
        mimeType: 'application/pdf'
      }
    })
  });

  if (response.getResponseCode() >= 300) {
    throw new Error('Document AI error ' + response.getResponseCode() + ': ' + response.getContentText().substring(0, 300));
  }

  var doc = JSON.parse(response.getContentText()).document;
  return parseInvoiceDocument_(doc);
}

function parseInvoiceDocument_(doc) {
  var entities = doc.entities || [];
  var ocrText = doc.text || '';

  function getEntity(type) {
    var found = entities.find(function(e) { return e.type === type; });
    if (!found) return '';
    return (found.normalizedValue && (found.normalizedValue.text || found.normalizedValue.dateValue)) || found.mentionText || '';
  }

  function getLineProp(lineItem, type) {
    var prop = (lineItem.properties || []).find(function(p) { return p.type === type; });
    if (!prop) return '';
    return prop.mentionText || (prop.normalizedValue && prop.normalizedValue.text) || '';
  }

  var lineItems = entities
    .filter(function(e) { return e.type === 'line_item'; })
    .map(function(lineItem) {
      var description = getLineProp(lineItem, 'line_item/description');
      var quantity = toNumber_(getLineProp(lineItem, 'line_item/quantity'));
      var unitPrice = toNumber_(getLineProp(lineItem, 'line_item/unit_price'));
      var amount = toNumber_(getLineProp(lineItem, 'line_item/amount'));
      var unit = standardizeUnit_(getLineProp(lineItem, 'line_item/unit'), description);

      if (!isValidLineItem_(description, quantity, amount)) return null;

      return {
        description: description,
        unit: unit,
        quantity: quantity,
        unitPrice: unitPrice,
        amount: amount
      };
    })
    .filter(function(x) { return x !== null; });

  return {
    vendor: getEntity('supplier_name') || getEntity('supplier') || 'Unknown Vendor',
    invoiceNumber: getEntity('invoice_id') || getEntity('invoice_number'),
    invoiceDate: getEntity('invoice_date'),
    projectOrCustomer: extractProjectOrCustomer_(ocrText),
    salesOrderOrPo: extractSalesOrderOrPo_(ocrText),
    lineItems: consolidateLineItems_(lineItems)
  };
}

function buildMasterRows_(parsed, sourceEmailId) {
  var timestamp = new Date().toISOString();
  var items = parsed.lineItems.length ? parsed.lineItems : [{}];
  var rows = [];

  for (var i = 0; i < items.length; i++) {
    var li = items[i];
    if (!li.description && !li.quantity && !li.amount) continue;

    rows.push([
      timestamp,
      parsed.vendor,
      parsed.invoiceNumber,
      parsed.invoiceDate,
      parsed.projectOrCustomer,
      parsed.salesOrderOrPo,
      i + 1,
      li.description || '',
      li.unit || '',
      li.quantity || '',
      li.unitPrice || '',
      li.amount || '',
      sourceEmailId || '',
      '',
      '',
      ''
    ]);
  }

  return rows;
}

function maybeForwardPdfsToAccounting_(processedPDFs) {
  if (!SETTINGS.ACCOUNTING_FORWARD_EMAIL) return;

  var invoiceList = processedPDFs.map(function(x) { return x.invoiceNumber; }).join(', ');

  if (SETTINGS.PAUSE_ACCOUNTING_EMAILS) {
    console.log('Accounting email paused. Would have forwarded: ' + invoiceList);
    return;
  }

  GmailApp.sendEmail(
    SETTINGS.ACCOUNTING_FORWARD_EMAIL,
    'Supplier Invoice Batch - ' + processedPDFs.length + ' invoices',
    'Invoices: ' + invoiceList,
    { attachments: processedPDFs.map(function(x) { return x.attachment; }) }
  );
}

function ensureMasterHeader_(sheet) {
  if (sheet.getLastRow() > 0) return;

  sheet.getRange(1, 1, 1, 16).setValues([[
    'ingested_at', 'vendor', 'invoice_number', 'invoice_date',
    'project_or_customer', 'sales_order_or_po', 'line_idx', 'description',
    'uom', 'qty', 'unit_price', 'pre_tax_amount', 'source_email_id',
    'material_type_final', 'review_status', 'review_note'
  ]]).setFontWeight('bold');
}

function buildInvoiceKeySet_(sheet) {
  var set = new Set();
  if (sheet.getLastRow() < 2) return set;

  sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues().forEach(function(row) {
    var vendor = (row[1] || '').toString().trim();
    var invoice = (row[2] || '').toString().trim();
    if (vendor && invoice) set.add(vendor + '|||' + invoice);
  });

  return set;
}

function getOrCreateLabel_(name) {
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}

function toNumber_(value) {
  if (value === null || value === undefined) return '';
  if (typeof value === 'number') return value;
  var n = parseFloat(String(value).replace(/[$,%\s,]/g, ''));
  return isNaN(n) ? '' : n;
}

function isValidLineItem_(description, quantity, amount) {
  if (!description) return false;

  var desc = description.toString().trim().toLowerCase();
  var invalidPatterns = [
    /terms?:/,
    /service charge/,
    /apr|interest/,
    /past due|balance/,
    /^page \d+/,
    /total|subtotal|tax/,
    /phone|fax|email/,
    /www\.|\.com/,
    /remit payment/
  ];

  if (invalidPatterns.some(function(pattern) { return pattern.test(desc); })) return false;
  if (!quantity && !amount) return false;
  if (desc.length < 3) return false;

  return true;
}

function standardizeUnit_(unit, description) {
  if (!unit) return '';

  var u = unit.toString().trim().toUpperCase();
  var desc = (description || '').toString().toUpperCase();

  if (u === 'PVC' || desc.match(/PVC\s+(90|45|22\.5|11\.25)/) || desc.match(/(ELBOW|FITTING|TEE|COUPLING)/)) {
    return 'Ea';
  }

  var map = {
    'EA': 'Ea',
    'EACH': 'Ea',
    'LF': 'Lf',
    'IF': 'Lf',
    'LIN FT': 'Lf',
    'LINEAR FEET': 'Lf',
    'FT': 'Lf',
    'FEET': 'Lf',
    'SQ FT': 'SF',
    'SQFT': 'SF',
    'GAL': 'Gal',
    'GALLON': 'Gal'
  };

  return map[u] || unit.toString().trim();
}

function consolidateLineItems_(items) {
  var grouped = {};

  (items || []).forEach(function(item) {
    if (!item.description) return;
    var key = item.description.toString().toUpperCase() + '|||' + (item.unit || '').toString().toUpperCase();

    if (!grouped[key]) {
      grouped[key] = Object.assign({}, item);
    } else {
      grouped[key].quantity = (grouped[key].quantity || 0) + (item.quantity || 0);
      grouped[key].amount = (grouped[key].amount || 0) + (item.amount || 0);
      if (item.unitPrice) grouped[key].unitPrice = item.unitPrice;
    }
  });

  return Object.values(grouped);
}

function extractProjectOrCustomer_(ocrText) {
  // Vendor-specific customization goes here.
  // Examples: PO number, job name, customer name, ship-to location, project code.
  return '';
}

function extractSalesOrderOrPo_(ocrText) {
  if (!ocrText) return '';

  var soMatch = ocrText.match(/Sales\s+Order\s+#?\s*([A-Z]*\d+)/i);
  if (soMatch) return soMatch[1].toString().toUpperCase();

  var poMatch = ocrText.match(/P\.?O\.?\s*(?:No\.?|#)?\s*([A-Z0-9\-]+)/i);
  if (poMatch) return poMatch[1].toString().toUpperCase();

  return '';
}
