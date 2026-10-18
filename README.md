# Chemistry Building Blocks v2

A historical VBA prototype for selecting an existing chemistry diagram from a
natural-language query and inserting its Word Quick Part. It selects a stored
building block; it does not generate a new diagram.

## Origin and relationship to the base toolkit

[Kuang Wen](https://github.com/kwen1510) created this repository and uploaded the
[VBA fragment on 28 January 2024](https://github.com/String-sg/chemistry-building-blocks-v2/commit/8451c9e8b6ccd1a636bd56b378c631e6f98a243b).
The [original description](https://github.com/String-sg/chemistry-building-blocks-v2/blob/db31f1a3236fbb62714c5ec5962eec37ad9131f9/README.md)
identifies it as an additional layer over Microsoft Word's Quick Parts in
[Chemistry Building Blocks](https://github.com/String-sg/Chemistry-Building-Blocks).

The base toolkit supplies the reusable Word diagram assets. This repository
contains only `Code.txt` and this README: it does not include those assets, a Word
template, a packaged add-in, or the complete host implementation. The `v2` name
identifies this experiment, rather than a complete replacement for the base
toolkit.

## What the prototype does

[`Code.txt`](Code.txt) contains `QueryDiagramLabel` and a small
`ExtractJsonValue` helper. The intended flow is:

1. Ask for a query using a VBA `InputBox`.
2. Send a prompt asking a service to choose from four diagram labels:
   `exothermic energy profile diagram`, `endothermic energy profile diagram`,
   `blotzmann distribution diagram` (spelled this way in the source), and
   `electrolytic cell diagram`.
3. Read a `response` value from the service's JSON reply.
4. If the lowercased output exactly matches `exothermic energy profile diagram`,
   call `InsertBuildingBlock "Exothermic"`. Every other output displays
   “There are no relevant diagrams”.

Only the exothermic label has an insertion branch. Listing four labels in the
prompt does not implement insertion for all four.

## Incomplete integration and compatibility

This is source for inspection and adaptation, not a ready-to-install extension.
Copying these files alongside a Word template does not install the prototype.
A developer would need to supply and validate:

- A host Word/VBA project with an `InsertBuildingBlock` implementation and an
  existing Quick Part named `Exothermic`. Neither is defined or included here.
- A reachable service URL in place of `apiUrl = "..."`. The source comment refers
  to an author's PaLM-based service; no service implementation or live endpoint
  is supplied, and that comment does not establish current provider support.
- A matching request/response contract. The fragment sends `query_string`,
  `temperature`, and `max_output_tokens` and expects a `response` field. It is not
  a complete adapter for a current model provider; request encoding, response
  parsing and failure handling need validation before use.
- A compatible execution environment. The fragment creates the Windows COM
  object `MSXML2.XMLHTTP`. This repository provides no verified modern Word/OS
  compatibility matrix or cross-platform implementation.

The code has not changed since January 2024; later commits through June 2025
changed the README. The repository does not demonstrate enhanced activity blocks,
new interface templates, or compatibility improvements. No automated tests or
Word execution results are included. Any live adaptation would send the entered
query to the configured external service.

## Status and possible future work

Retain this as a historical prototype with its attribution. There is no supported
installation or product release here. Consolidation into the base toolkit's
historical examples, or archival of this repository, requires a separate owner
decision.

A bounded product hypothesis is natural-language retrieval of vetted, editable
chemistry diagrams. Its value is unvalidated: compare it with manual selection
and a searchable labelled catalogue using fixed teacher queries, selection
accuracy and preparation time. Revival would need a maintainer, demonstrated
teacher demand, complete label mappings, an explicit host/service contract and
tested Word/OS versions before making installation or compatibility claims.
