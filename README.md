# SAP Production Dosing Automation

A sanitized portfolio project documenting an Excel VBA workflow that prepares SAP production-order data, records status in Excel, and attaches dosing-order PDF documents.

> The original production workbook is intentionally excluded. This public repository contains documentation and synthetic examples only.

## Project overview

The reviewed application uses Excel as an operator-friendly input grid and SAP GUI scripting as the automation layer. It supports three connected activities:

1. Create and save unique SAP production orders from Excel rows.
2. Prepare confirmation and goods-movement data for review before optional posting.
3. Find matching dosing PDFs and attach them to SAP orders through the SAP GOS attachment flow.

The workbook writes status and timestamps back to Excel and records failures in an Exception_Log worksheet.

## Implemented features

- Header-based input validation and dynamic SAP status columns.
- Positive-value checks for yield quantity and final Brix.
- Duplicate-order detection and consistency validation.
- One-time processing of repeated order rows in the order-creation flow.
- SAP GUI session discovery, control checks, status-bar checks, and timeout protection.
- Production-order creation with material, plant, order type, quantity, dates, batch, and Brix mapping.
- Confirmation/goods-movement preparation for contiguous order groups.
- Component entry for raw material, quantity, storage location, and batch.
- Review checkpoint with automatic SAP posting disabled in the reviewed configuration.
- Three-stage PDF matching: exact prefix, normalized filename comparison, and a controlled suffix fallback.
- Multiple-PDF attachment support with duplicate suppression.
- Row-level statuses, timestamps, colored attachment results, message boxes, and Exception_Log entries.

The workbook uses two worksheet shape buttons. It does not contain VBA UserForms. The public version does not include the original VBA source because it contains environment-specific SAP identifiers and internal paths.

## Workflow

1. Populate INPUT_SHEET with one or more order rows. An order may span multiple component rows.
2. Keep rows for the same order contiguous for the confirmation flow.
3. Open SAP GUI and sign in to an authorized environment.
4. Run the main worksheet action. Process 1 validates data and creates each unique production order.
5. Process 2 replaces goods-movement rows and enters component data. The reviewed configuration writes READY FOR REVIEW - not posted.
6. Review SAP status, Excel timestamps, and Exception_Log.
7. Run the attachment action. Process 3 searches the configured PDF folder, attaches every unique match, and records the result.
8. Reconcile the Excel output against SAP.

## Input schema

| Column | Field | Use |
|---|---|---|
| A | Order Number | All three processes |
| B | Finished Material | Production-order creation |
| C | Raw Material | Goods movements |
| D | Batch Number | Goods movements |
| E | Quantity | Goods movements; Raw Quantity is also accepted |
| F | Unit | Operator context |
| G | Yield Quantity | Production-order creation |
| H | Final Brix % | Production-order creation |
| I | SAP Status | Process result |
| J | SAP Timestamp | Process result time |
| M | Attachment Status | PDF and SAP attachment result |
| N | Attachment Timestamp | Attachment result time |

The reviewed configuration uses the computer date for SAP basic start and end dates. Columns K and L are not used by the active code.

## Synthetic example

~~~csv
Order Number,Finished Material,Raw Material,Batch Number,Quantity,Unit,Yield Quantity,Final Brix %
ORD-DEMO-001,FG-DEMO-100,RM-DEMO-001,BATCH-DEMO-01,7200,KG,11000,66.5
ORD-DEMO-001,FG-DEMO-100,RM-DEMO-002,BATCH-DEMO-02,2.4,L,11000,66.5
~~~

These values are for documentation and mock testing only. They are not valid production identifiers.

## Technical details

- Language: VBA for Microsoft Excel.
- Workbook type: Excel macro-enabled workbook.
- VBA structure: one standard module plus ThisWorkbook and Sheet1 document modules.
- Interface: worksheet shape buttons, status cells, timestamps, and message boxes.
- Integration: late-bound SAP GUI COM automation through GetObject("SAPGUI") and GetScriptingEngine.
- Supporting runtime objects: standard VBA, Excel object model, Scripting.Dictionary, Dir, and file-attribute checks.
- No formulas, Excel Tables, PivotTables, charts, Power Query objects, database connectors, or external workbook links were detected in the reviewed package.

### Important configuration behavior

| Setting | Reviewed behavior | Public treatment |
|---|---|---|
| Current-date mode | Uses the computer date | Make configurable or mock it |
| SAP posting flag | False; review-only status | Keep disabled in demos |
| Plant/order/storage values | Environment-specific constants | Externalize or replace with placeholders |
| PDF folder | Internal network location | Inject privately at runtime |
| SAP control IDs | Tied to one SAP screen layout | Keep private or mock |

## Requirements for private internal use

- Windows desktop Excel with VBA support.
- SAP GUI for Windows with SAP GUI scripting enabled.
- Authorized SAP access for production-order, confirmation, goods-movement, and attachment actions.
- Approved access to the PDF document location.
- A trusted, private copy of the workbook.

The public repository does not connect to SAP and does not require Python, a database, or a web API.

## Setup

### Private/internal deployment

1. Keep the original workbook in an approved internal location.
2. Configure environment-specific SAP settings and the private PDF location in a controlled copy.
3. Open SAP GUI and sign in.
4. Populate INPUT_SHEET with the required headers and operational data.
5. Run the main process and review status columns and Exception_Log.
6. Run the attachment process after confirming the PDF matching rules.
7. Reconcile results in SAP before closing the session.

### Public demonstration

Use the synthetic CSV and a mock SAP adapter. Keep automatic posting disabled and use fake PDF filenames. Do not connect the public repository to a live SAP system.

## Known limitations

- The confirmation process expects rows for the same order to be adjacent.
- An order group with more than six component rows is rejected.
- The suffix-based PDF fallback can produce false matches if filenames are not controlled.
- SAP GUI control IDs can change when the SAP screen layout changes.
- A failed validation stops the main process; attachment failures can be logged and skipped per order where possible.
- The original implementation uses hard-coded environment settings that must be externalized for reuse.

## Recommended repository structure

~~~text
.
├── README.md
├── docs/
│   ├── architecture.md
│   └── redaction-checklist.md
├── examples/
│   └── input-template.csv
├── src/
│   └── vba/
│       └── README.md
└── .gitignore
~~~

## Screenshots to add later

Use only synthetic or fully masked data:

1. Input sheet with two component rows and the two action buttons.
2. Review-mode status and timestamp columns.
3. Successful and missing-PDF attachment results.
4. Synthetic Exception_Log entry.
5. Architecture diagram showing Excel, VBA, SAP GUI scripting, PDF matching, and Excel logging.

## Business benefits

This automation reduces repetitive SAP entry, standardizes field mapping, prevents duplicate order creation, adds validation before SAP changes, creates a review checkpoint before posting, automates document attachment, and gives operators a traceable status and exception history.

## Implemented versus future improvements

Implemented: validation, duplicate handling, SAP GUI automation, review-mode confirmation preparation, PDF matching, attachment processing, timestamps, and exception logging.

Future improvements: configuration externalization, a mock SAP adapter, automated tests, deterministic PDF-to-order mapping, explicit approval before posting, structured audit retention, and a controlled business-date input.

## Publication and security notes

Do not publish the original workbook, compiled vbaProject.bin, production xlsm, real order/material/batch values, exception history, SAP control mappings, internal SharePoint or UNC paths, credentials, or live screenshots. Review workbook metadata and Git history for secrets before adding any future source exports.

No plaintext password was found during static review, but the original embedded VBA project contains protection metadata and should remain private unless the owner approves release.

## Resume description

> Developed an Excel VBA automation tool that integrates with SAP GUI to create production orders, prepare confirmation and goods-movement data for review, and attach dosing documents. Implemented input validation, duplicate-order handling, status and timestamp tracking, network PDF matching, SAP error detection, and exception logging to improve process traceability and reduce manual operations effort.

## License and ownership

Choose a license only after confirming that the workbook code, SAP integration logic, process knowledge, and sample data can be published. For employer-owned work, keep the original implementation private and publish only a sanitized demonstration.
