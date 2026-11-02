# Architecture

The portfolio version documents a three-stage Excel-to-SAP workflow:

~~~text
Excel input rows
      |
      v
VBA validation and grouping
      |
      +--> SAP GUI scripting: create production order
      |
      +--> SAP GUI scripting: prepare confirmation and goods movements
      |
      +--> PDF filename matching --> SAP GOS attachment
      |
      v
Excel statuses, timestamps, and exception log
~~~

## Main components

- Excel worksheet front end for inputs, action buttons, statuses, and timestamps.
- VBA orchestration layer for validation, grouping, SAP actions, and error handling.
- SAP GUI scripting adapter using late-bound COM automation.
- PDF matching layer using exact, normalized, and fallback filename rules.
- Excel-based operational logging through statuses and Exception_Log.

## Public-demo boundary

The public repository contains documentation and synthetic examples only. The production workbook, compiled VBA project, SAP control IDs, and live network paths are deliberately excluded. A future runnable demo should use a mock SAP adapter so the workflow can be tested without changing live business data.
