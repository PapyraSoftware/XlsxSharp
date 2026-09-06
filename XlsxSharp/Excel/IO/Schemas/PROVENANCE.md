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
Open Specification Promise).

Not every part kind is validated: legacy VML (`xl/drawings/vmldrawing*.vml`) and
`docProps/core.xml` are skipped deliberately (see `SchemaValidator.cs`) rather than
pulling in schemas with far larger transitive closures (the full WordprocessingML
schema tree, external Dublin Core schemas) for parts that are small and rarely wrong.
