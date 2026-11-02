# VBA source boundary

The original VBA project was reviewed for this portfolio README but is intentionally not included in the public repository.

Reasons for exclusion:

- It contains environment-specific SAP screen identifiers and configuration.
- It contains an internal network location for PDF documents.
- The source workbook contains operational production data and exception history.
- The compiled macro project contains document and protection metadata that requires owner approval before release.

For a future public code release, export only a reviewed and sanitized version of the standard module and document modules. Replace live SAP calls with a mock adapter, externalize configuration, keep automatic posting disabled, and test the sanitized code with synthetic identifiers.
