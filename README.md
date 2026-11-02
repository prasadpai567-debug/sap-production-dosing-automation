# SAP Production Dosing Automation

A sanitized portfolio project documenting an Excel VBA and SAP GUI Scripting workflow for preparing production-order dosing data, staging confirmation and goods-movement information for review, and attaching dosing-related PDF documents.

> **Public status:** This repository contains documentation and synthetic input data only. The production workbook, VBA source, SAP identifiers, internal paths, and operational records are intentionally excluded. The public repository is not runnable against SAP.

## Business problem

The documented private workflow addresses a repetitive production-operations process:

1. Read production-order and component information from an Excel input grid.
2. Enter or prepare corresponding data in SAP.
3. Locate dosing-related PDF documents.
4. Attach documents to the relevant SAP orders.
5. Reconcile SAP results and maintain an operational exception history.

Without structured automation, these steps can require repeated data entry, manual checking, filename searching, and status reconciliation.

## How the documented automation helps

The private workbook is designed around Excel as an operator-friendly input and review surface, with SAP GUI Scripting used for environment-specific SAP interaction. It validates input before SAP changes, groups related component rows, records status and timestamps, and supports PDF matching and attachment.

The public repository documents this workflow but does not include the executable workbook or a live SAP connection.

## Public verification status

| Area | Public artifact | Verification status |
|---|---|---|
| Workflow and business context | README.md, docs/architecture.md | Inspectable documentation |
| Synthetic input format | examples/input-template.csv | Inspectable example data |
| Production workbook | Not included | Not publicly verifiable |
| VBA source | Not included | Not publicly verifiable |
| SAP GUI interaction | Not included | Not publicly testable |
| Mock adapter and demo | Not included | Planned future improvement |
| Automated tests | Not included | No automated test evidence |

The feature list below describes behavior documented from the private workbook. It should not be interpreted as executable functionality present in this public repository.

## Documented private-workbook behavior

- Header-based input validation and dynamic status columns.
- Positive-value checks for yield quantity and final Brix.
- Duplicate-order detection and consistency validation.
- One-time processing of repeated order rows during order creation.
- SAP GUI session discovery, control checks, status-bar checks, and timeout protection.
- Production-order creation with material, plant, order type, quantity, dates, batch, and Brix mapping.
- Confirmation and goods-movement preparation for contiguous order groups.
- Component entry for raw material, quantity, storage location, and batch.
- Review checkpoint with automatic SAP posting disabled in the reviewed configuration.
- Three-stage PDF matching: exact prefix, normalized filename comparison, and controlled suffix fallback.
- Multiple-PDF attachment support with duplicate suppression.
- Row-level statuses, timestamps, colored attachment results, message boxes, and Exception_Log entries.

## Workflow

1. Populate the private workbook input sheet with one or more order rows. An order may span multiple component rows.
2. Keep rows for the same order contiguous for the confirmation flow.
3. Open SAP GUI and sign in to an authorized environment.
4. Run the main worksheet action. The documented first process validates data and creates each unique production order.
5. Prepare confirmation and goods-movement data. The reviewed configuration writes READY FOR REVIEW rather than posting automatically.
6. Review SAP status, Excel timestamps, and Exception_Log.
7. Run the attachment action. The documented third process searches the configured PDF folder and records the result.
8. Reconcile the Excel output against SAP.

## Input format

The public CSV is a synthetic, input-only example:

| CSV column | Field | Documented use |
|---|---|---|
| A | Order Number | Order grouping and processing |
| B | Finished Material | Production-order preparation |
| C | Raw Material | Goods-movement preparation |
| D | Batch Number | Goods-movement preparation |
| E | Quantity | Goods-movement preparation |
| F | Unit | Operator context |
| G | Yield Quantity | Production-order preparation |
| H | Final Brix % | Production-order preparation |

The private workbook also documents output/status columns such as SAP status, SAP timestamp, attachment status, and attachment timestamp. Those columns are not included in the public CSV template.

### Synthetic example

~~~csv
Order Number,Finished Material,Raw Material,Batch Number,Quantity,Unit,Yield Quantity,Final Brix %
ORD-DEMO-001,FG-DEMO-100,RM-DEMO-001,BATCH-DEMO-01,7200,KG,11000,66.5
ORD-DEMO-001,FG-DEMO-100,RM-DEMO-002,BATCH-DEMO-02,2.4,L,11000,66.5
~~~

These values are synthetic and are not valid production identifiers.

## Technologies

- Microsoft Excel and Excel VBA.
- SAP GUI Scripting through late-bound COM automation.
- Excel worksheet controls, cells, timestamps, and message boxes.
- Local or network PDF filename matching.
- Excel-based status and exception logging.

The public repository does not require Python, a database, or a web API.

## Private/internal prerequisites

- Windows desktop Excel with VBA support.
- SAP GUI for Windows with SAP GUI Scripting enabled.
- Authorized access to the relevant SAP environment.
- Approved access to the private PDF document location.
- A trusted private copy of the original workbook.

## Public usage

The current public repository has no executable command or runnable demo.

You can inspect the synthetic input template and the architecture documentation. Do not connect the public repository to a live SAP system. A future public demonstration should use a mock SAP adapter, synthetic identifiers, fake PDF filenames, and automatic posting disabled.

## Error handling and logging

The documented private workflow includes:

- Input validation before SAP actions.
- SAP session and control checks.
- SAP status-bar checks and timeout protection.
- Row-level status and timestamp updates.
- An Exception_Log worksheet.
- Attachment result recording and duplicate suppression.

Because the VBA source is not public, these behaviors have not been independently executed or tested from this repository.

## Business benefits

The documented workflow is intended to:

- Reduce repetitive SAP data entry.
- Standardize field mapping from Excel to SAP.
- Prevent duplicate order processing.
- Provide a review checkpoint before posting.
- Automate document matching and attachment.
- Improve status visibility and exception traceability.

No measured time-saving or productivity metric is published because no validated measurement is included in this repository.

## Known limitations

- The confirmation process expects rows for the same order to be adjacent.
- An order group with more than six component rows is rejected in the documented configuration.
- Suffix-based PDF fallback matching can produce false matches if filenames are not controlled.
- SAP GUI control IDs can change when the SAP screen layout changes.
- A failed validation stops the main process; attachment failures may be logged and skipped per order where possible.
- The original workbook uses environment-specific settings that must be externalized for reuse.
- The public repository cannot be used to execute the documented SAP workflow.

## Planned improvements

- Add a sanitized mock adapter and a runnable public demonstration.
- Add automated tests for validation, grouping, duplicate handling, PDF matching, and failure paths.
- Externalize configuration and support an explicit business date.
- Replace fallback matching with deterministic PDF-to-order mapping.
- Add structured audit retention and controlled approval before any posting.
- Add synthetic screenshots and an architecture diagram.
- Add continuous integration after executable public code is introduced.

## Screenshots and demonstration assets

Future screenshots should use only synthetic or fully masked data:

1. Input sheet with two component rows and the worksheet action buttons.
2. Review-mode status and timestamp columns.
3. Successful and missing-PDF attachment results.
4. Synthetic Exception_Log entry.
5. Architecture diagram showing Excel, VBA, SAP GUI Scripting, PDF matching, and Excel logging.

## Repository structure

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

## Security and publication boundary

Do not publish the original workbook, compiled vbaProject.bin, production .xlsm, real order/material/batch values, exception history, SAP control mappings, internal SharePoint or UNC paths, credentials, or live screenshots.

Before publishing any future source export or workbook:

1. Remove company, personal, and author metadata.
2. Remove local, SharePoint, OneDrive, and UNC paths.
3. Replace operational identifiers with synthetic values.
4. Inspect the ZIP contents of any .xlsm.
5. Review Git history for secrets that may have been committed and later deleted.
6. Confirm that ownership and publication rights are clear.

The safe public default is documentation plus synthetic examples. Keep the original macro workbook in a private, approved location.

## Recruiter summary

> Developed and documented an Excel VBA and SAP GUI Scripting workflow for production-order dosing operations. The workflow validates production data, groups component rows, prepares SAP confirmation and goods-movement information for review, matches dosing PDFs, and records statuses, timestamps, and exceptions. The public repository demonstrates the process architecture and security-conscious publication approach using synthetic data while keeping the production workbook and environment-specific SAP configuration private.

## License and ownership

No license is currently included. Add one only after confirming that the workbook code, SAP integration logic, process knowledge, and sample data can be published. For employer-owned work, keep the original implementation private and publish only a sanitized demonstration.
