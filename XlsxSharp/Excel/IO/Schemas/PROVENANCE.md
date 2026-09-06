# Schema provenance

The `.xsd` files in this directory are the OOXML/OPC schemas `SaveOptions.ValidatePackage`
validates saved parts against (see `SchemaValidator.cs`).

- `sml.xsd`, `dml-*.xsd`, `shared-*.xsd` are from **ISO/IEC 29500-4:2016**
  ("Information technology — Document description and processing languages — Office Open
  XML File Formats — Part 4: Transitional Migration Features"), the SpreadsheetML and
  DrawingML schemas.
- `opc-contentTypes.xsd`, `opc-relationships.xsd` are from **ECMA-376** (the OPC —
  Open Packaging Conventions — schemas for `[Content_Types].xml` and `.rels` parts).

Both standards are published by Ecma International, which permits free copying and
distribution of its standards, and the identical content is also published as
ISO/IEC 29500 (freely available in this form via Microsoft's implementation of the
Open Specification Promise). These particular copies were taken from a schema bundle
already present in the sandbox this change was authored in rather than downloaded fresh
from ecma-international.org (blocked by that session's network policy) - their exact
distribution chain before that point was not independently re-verified, so if that
matters for your use of this project, cross-check against a copy fetched directly from
Ecma International or ISO.

Not every part kind is validated: legacy VML (`xl/drawings/vmldrawing*.vml`) and
`docProps/core.xml` are skipped deliberately (see `SchemaValidator.cs`) rather than
pulling in schemas with far larger transitive closures (the full WordprocessingML
schema tree, external Dublin Core schemas) for parts that are small and rarely wrong.
