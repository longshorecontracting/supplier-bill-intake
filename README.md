# Mechanical Insulation Supplier Bill Automation — LLM Build Brief

Use this document to help an LLM build a version of the workflow for **your own company**, vendors, Google account, and accounting system.

This is based on a working mechanical insulation supplier-bill workflow, but all live company-specific IDs, URLs, labels, and accounting endpoints should be replaced before use.

---

## 1. Plain-English Goal

I run a mechanical insulation company. I receive PDF supplier bills/invoices by email. I want an automation that:

1. Watches a Gmail label for vendor invoice emails.
2. Finds PDF attachments.
3. Skips non-invoice documents such as statements, packing slips, delivery receipts, and credit memos.
4. Sends each invoice PDF through Google Document AI or another invoice-parsing/OCR service.
5. Extracts invoice header fields and line items.
6. Writes one row per line item into a Google Sheets `Master` tracker.
7. Avoids duplicate invoice imports.
8. Moves successfully processed emails to a filed label.
9. Sends failures to an error label.
10. Optionally forwards the original PDFs to QuickBooks Online or another accounting inbox.
11. Optionally builds a pricing catalog from the line-item history.

---

## 2. Recommended System Architecture

### Required tools

- Gmail
- Google Sheets
- Google Apps Script
- Google Document AI Invoice Parser, or another invoice extraction service
- Optional: QuickBooks Online bill forwarding or QBO API

### Core tabs in Google Sheets

#### `Master`

This is the main line-item tracker.

Recommended columns:

| Column | Field |
|---|---|
| A | ingested_at |
| B | vendor |
| C | invoice_number |
| D | invoice_date |
| E | project_or_customer |
| F | sales_order_or_po |
| G | line_idx |
| H | description |
| I | uom |
| J | qty |
| K | unit_price |
| L | pre_tax_amount |
| M | source_email_id |
| N | material_type_final |
| O | review_status |
| P | review_note |

#### `Pricing Catalog` optional

Generated from `Master`. Shows latest price, prior price, percent change, times purchased, first seen, last seen, and price history.

#### `Catalog Aliases` optional

Manual cleanup table for merging duplicate item names caused by OCR/vendor description differences.

---

## 3. Setup Questions to Give the LLM

Before writing code, ask me these questions:

1. What Gmail label should contain invoices waiting to process?
2. What Gmail label should processed invoices move to?
3. What Gmail label should failed invoices move to?
4. Which vendors should be supported first?
5. Do the vendors send only invoices, or also statements/packing slips/credits?
6. What Google Sheet should receive the line items?
7. Do I want to forward processed PDFs to QuickBooks Online?
8. Do I need to map invoices to projects from QuickBooks, or will the project be on the invoice/P.O. field?
9. What material categories do I want to classify line items into?
10. Do I need a price catalog, or just the raw master tracker?

---

## 4. Sanitized Starter Code — Invoice Intake Pipeline

Paste this into Google Apps Script and have the LLM adapt it.

```javascript
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
```

---

## 5. Sanitized Starter Code — Optional Pricing Catalog

Use this only after the `Master` tab is populated.

```javascript
// =================================================================================================
// OPTIONAL PRICING CATALOG — SANITIZED TEMPLATE
// Purpose:
//   Reads Master line-item history and creates a Pricing Catalog tab with latest price,
//   prior price, price change %, purchase count, and compact price history.
// =================================================================================================

var PC = {
  SPREADSHEET_ID: 'REPLACE_WITH_GOOGLE_SHEET_ID',
  MASTER_SHEET:   'Master',
  CATALOG_SHEET:  'Pricing Catalog',
  ALIAS_SHEET:    'Catalog Aliases'
};

var SITE_PREFIXES = [
  // Add vendor/location prefixes to strip from descriptions.
  // Examples: 'W JAX', 'W RAL', 'TAMPA', 'ORLANDO'
];

var MATERIAL_LABELS = {
  rubberpipe:      'Rubber – Pipe',
  rubbersheet:     'Rubber – Sheet',
  fiberglasspipe:  'Fiberglass – Pipe',
  ductwrap:        'Fiberglass – Duct Wrap',
  ductboard:       'Fiberglass – Duct Board',
  tankwrap:        'Tank Wrap',
  pvcfitting:      'PVC – Fittings',
  pvcstraight:     'PVC – Straight Jacket',
  pvcsheet:        'PVC – Sheet',
  polyisopipe:     'Polyiso – Pipe',
  polyisofitting:  'Polyiso – Fittings',
  foamglaspipe:    'Foamglas – Pipe',
  foamglasfitting: 'Foamglas – Fittings',
  metallf:         'Metal – Straight LF',
  metalsheet:      'Metal – Sheet',
  metalfitting:    'Metal – Fittings',
  adhesive:        'Adhesives & Coatings',
  mastic:          'Mastics & Coatings',
  tape:            'Tapes',
  accessory:       'Accessories',
  other:           'Other'
};

function buildPricingCatalog() {
  var ss = SpreadsheetApp.openById(PC.SPREADSHEET_ID);
  var master = ss.getSheetByName(PC.MASTER_SHEET);
  if (!master) throw new Error('Master sheet not found.');

  var lastRow = master.getLastRow();
  if (lastRow < 2) {
    console.log('Master sheet is empty.');
    return;
  }

  ensureAliasSheet_(ss);
  var aliases = loadAliases_(ss);

  // Expected Master columns:
  // A ingested_at, B vendor, C invoice_number, D invoice_date,
  // E project_or_customer, F sales_order_or_po, G line_idx, H description,
  // I uom, J qty, K unit_price, L pre_tax_amount, M source_email_id,
  // N material_type_final, O review_status, P review_note
  var data = master.getRange(2, 1, lastRow - 1, 16).getValues();
  var groups = {};

  data.forEach(function(row) {
    var invoiceDate = row[3];
    var invoiceNumber = (row[2] || '').toString().trim();
    var rawDescription = (row[7] || '').toString();
    var uom = (row[8] || '').toString().trim();
    var unitPrice = row[10];
    var materialType = (row[13] || 'other').toString().trim();

    if (!rawDescription || !unitPrice || typeof unitPrice !== 'number' || unitPrice <= 0) return;

    var cleanDescription = normalizeDescription_(rawDescription);
    var canonicalDescription = applyAlias_(cleanDescription, aliases);
    var key = materialType + '|||' + canonicalDescription.toUpperCase() + '|||' + uom.toUpperCase();

    if (!groups[key]) {
      groups[key] = {
        material_type: materialType,
        description: canonicalDescription,
        uom: uom,
        occurrences: []
      };
    }

    groups[key].occurrences.push({
      unit_price: unitPrice,
      invoice_date: invoiceDate instanceof Date ? invoiceDate : new Date(invoiceDate),
      invoice_number: invoiceNumber
    });
  });

  var tz = Session.getScriptTimeZone();
  var catalogRows = [];

  Object.keys(groups).forEach(function(key) {
    var group = groups[key];
    var occ = group.occurrences.sort(function(a, b) {
      return a.invoice_date - b.invoice_date;
    });

    var latest = occ[occ.length - 1];
    var previous = occ.length > 1 ? occ[occ.length - 2] : null;
    var latestPrice = latest.unit_price;
    var previousPrice = previous ? previous.unit_price : '';
    var changePct = previous && previous.unit_price > 0
      ? (latestPrice - previous.unit_price) / previous.unit_price
      : '';

    catalogRows.push([
      MATERIAL_LABELS[group.material_type] || group.material_type,
      group.description,
      group.uom,
      latestPrice,
      previousPrice,
      changePct,
      occ.length,
      formatDate_(occ[0].invoice_date, tz),
      formatDate_(latest.invoice_date, tz),
      latest.invoice_number,
      buildHistoryString_(occ, tz)
    ]);
  });

  catalogRows.sort(function(a, b) {
    return (a[0] + a[1]).localeCompare(b[0] + b[1]);
  });

  writePricingCatalog_(ss, catalogRows);
}

function writePricingCatalog_(ss, rows) {
  var sheet = ss.getSheetByName(PC.CATALOG_SHEET) || ss.insertSheet(PC.CATALOG_SHEET);
  sheet.clearContents();
  sheet.clearFormats();

  var headers = [
    'Material Type', 'Description', 'UOM', 'Latest Price', 'Prev Price',
    'Change %', 'Times Purchased', 'First Seen', 'Last Seen',
    'Last Invoice #', 'Price History'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    sheet.getRange(2, 4, rows.length, 2).setNumberFormat('$#,##0.0000');
    sheet.getRange(2, 6, rows.length, 1).setNumberFormat('+0.0%;-0.0%;–');
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, rows.length + 1, headers.length).createFilter();
  }
}

function ensureAliasSheet_(ss) {
  if (ss.getSheetByName(PC.ALIAS_SHEET)) return;

  var sheet = ss.insertSheet(PC.ALIAS_SHEET);
  sheet.getRange(1, 1, 1, 3).setValues([[
    'Variant normalized description contains...',
    'Canonical Name',
    'Notes'
  ]]).setFontWeight('bold');

  sheet.setFrozenRows(1);
}

function loadAliases_(ss) {
  var sheet = ss.getSheetByName(PC.ALIAS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];

  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues()
    .filter(function(row) { return row[0] && row[1]; })
    .map(function(row) {
      return {
        variant: row[0].toString().trim().toUpperCase(),
        canonical: row[1].toString().trim()
      };
    });
}

function applyAlias_(description, aliases) {
  var upper = description.toUpperCase();
  for (var i = 0; i < aliases.length; i++) {
    if (upper.indexOf(aliases[i].variant) !== -1) return aliases[i].canonical;
  }
  return description;
}

function normalizeDescription_(raw) {
  var s = raw.replace(/[\r\n\t]+/g, ' ').replace(/\s{2,}/g, ' ').trim();

  SITE_PREFIXES.forEach(function(prefix) {
    var p = prefix.toUpperCase();
    if (s.toUpperCase().indexOf(p + ' ') === 0) {
      s = s.substring(prefix.length + 1).trim();
    }
    if (s.toUpperCase().slice(-(p.length + 1)) === ' ' + p) {
      s = s.slice(0, -(prefix.length + 1)).trim();
    }
  });

  return s.replace(/[^\w\s\-\/\.\#\"\'\(\)\\,\+]/g, ' ').replace(/\s{2,}/g, ' ').trim();
}

function buildHistoryString_(occurrences, tz) {
  var deduped = [];
  var lastPrice = null;

  occurrences.forEach(function(o) {
    if (o.unit_price !== lastPrice) {
      deduped.push(o);
      lastPrice = o.unit_price;
    }
  });

  return deduped.map(function(o) {
    var dateLabel = o.invoice_date ? Utilities.formatDate(o.invoice_date, tz, "MMM ''yy") : '?';
    return '$' + o.unit_price.toFixed(4) + ' (' + dateLabel + ')';
  }).join(' → ');
}

function formatDate_(date, tz) {
  if (!date) return '';
  return Utilities.formatDate(date, tz, 'MM/dd/yyyy');
}
```

---

## 6. Prompt to Give an LLM

Copy/paste this prompt into ChatGPT, Claude, or Gemini:

```text
I run a mechanical insulation company and want to build a Google Apps Script automation for supplier invoice PDFs.

Goal:
- Watch a Gmail label for vendor invoice PDFs.
- Parse invoice PDFs using Google Document AI Invoice Parser.
- Write one row per invoice line item into a Google Sheets Master tab.
- Avoid duplicate imports by vendor + invoice number.
- Skip statements, packing slips, delivery receipts, and credit memos.
- Move successful emails to a filed label.
- Move failed emails to an error label.
- Optionally forward processed invoice PDFs to my accounting inbox or QuickBooks Online.
- Optionally create a Pricing Catalog tab from the Master line-item history.

My environment:
- Gmail label for new invoices: [INSERT LABEL]
- Processed label: [INSERT LABEL]
- Error label: [INSERT LABEL]
- Google Sheet ID: [INSERT SHEET ID]
- Vendor names: [INSERT VENDORS]
- Accounting forward email, if any: [INSERT EMAIL OR NONE]
- Do I use QuickBooks Online? [YES/NO]
- Do I want project mapping from QBO? [YES/NO]
- My common material categories are: [INSERT MATERIAL CATEGORIES]

Please help me:
1. Review the starter code.
2. Replace placeholders with my setup.
3. Identify what Google Cloud permissions/scopes I need.
4. Add test mode so no emails are moved or forwarded until I approve it.
5. Add logging to a ProcessingLog sheet.
6. Add a simple setup checklist for a non-programmer.
7. Explain where I should customize vendor-specific extraction rules.
```

---

## 7. Implementation Checklist

### Phase 1 — Build the sheet

- Create a Google Sheet.
- Add a `Master` tab.
- Let the script create headers automatically, or manually add the columns listed above.

### Phase 2 — Prepare Gmail

Create labels like:

- `Suppliers/VendorName/To-Process`
- `Suppliers/VendorName/To-File`
- `Suppliers/VendorName/Errors`

Create a Gmail filter that sends invoices from the vendor into the `To-Process` label.

### Phase 3 — Prepare Document AI

- Create a Google Cloud project.
- Enable Document AI API.
- Create an Invoice Parser processor.
- Copy the processor endpoint into `SETTINGS.PREDICTION_ENDPOINT`.
- Make sure Apps Script has authorization to call it.

### Phase 4 — Test safely

Keep these settings during testing:

```javascript
ACCOUNTING_FORWARD_EMAIL: '',
PAUSE_ACCOUNTING_EMAILS: true
```

Start with one or two known invoices.

Confirm:

- Invoice number is correct.
- Invoice date is correct.
- Vendor is correct.
- Quantities are correct.
- Unit prices are correct.
- Amounts are correct.
- Non-invoice attachments are skipped.
- Duplicate invoices are not imported twice.

### Phase 5 — Add optional features

Add only after the basic import works:

- Pricing Catalog
- Material classification
- Review queue
- QBO project mapping
- QBO forwarding
- Nightly time trigger

---

## 8. What Not to Copy Blindly

Do not copy another company’s live values for:

- Google Cloud project ID
- Document AI processor ID
- Google Sheet ID
- Web app URL
- QuickBooks/Intuit tokens
- QuickBooks company ID / realm ID
- Accounting forwarding email
- Gmail labels
- Vendor-specific regex patterns
- Internal project names

---

## 9. Suggested Build Order

Do not try to build the whole system at once.

Recommended order:

1. Gmail label → PDF detection.
2. PDF → Document AI response.
3. Document AI response → `Master` rows.
4. Duplicate prevention.
5. Move processed/error emails.
6. Add accounting forwarding.
7. Add pricing catalog.
8. Add material classification.
9. Add QBO project mapping.
10. Add nightly trigger.

---

## 10. Good First Test Case

Use one simple supplier invoice PDF with 3–10 line items.

Expected result:

- One invoice becomes multiple rows in `Master`.
- Header fields are repeated on each line.
- Each line has description, UOM, qty, unit price, and amount.
- The email is moved only after successful import.
- Running the script a second time does not duplicate the invoice.

---

## 11. Notes for Mechanical Insulation Vendors

Common things the LLM should account for:

- Fiberglass pipe insulation may be sold by linear foot, carton, or each.
- PVC fittings are often each, even when OCR misreads the unit.
- ASJ/FSK tape and jacketing may have inconsistent descriptions.
- Vendor branch prefixes may appear in item descriptions.
- Statements and packing slips may look like invoices but should not enter the tracker.
- Freight, tax, service charge, finance charge, and remittance text should usually be excluded from material unit pricing.
- A review queue is useful because OCR and vendor descriptions will not be perfect.

---

## 12. Recommended Next Improvement

After the basic tracker works, build a `ReviewQueue` tab that catches rows where:

- Material type is blank.
- Unit price is missing.
- Quantity is missing.
- Description looks suspicious.
- OCR confidence is low, if confidence scores are available.
- The same product appears under multiple similar names.

Then let the user approve/correct rows before they become final pricing catalog data.
