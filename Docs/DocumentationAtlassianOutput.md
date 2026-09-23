# Atlassian (Confluence) documentation output

[`Internal/Documentation/OutputProviders/DocumentationOutputAtlassian.ps1`](../Internal/Documentation/OutputProviders/DocumentationOutputAtlassian.ps1)

Emits Confluence **storage format** (XHTML plus `<ac:*>` macros) for a documentation
run: paste it into a page's *Rich Text -> Source* view, or POST it to the Confluence
REST API as `representation=storage`. Structurally it mirrors the HTML provider
(BasicInfo / FilteredSettings / ComplianceActions / ApplicabilityRules / Assignments /
CustomTables plus a per-run table of contents); only the markup differs.

Select it with `-OutputFormat atlassian`.

| Option (`$Options.Outputs.atlassian`) | Default | Meaning |
| --- | --- | --- |
| `AtlassianDocumentName` | `%MyDocuments%\%Organization%-%Date%.html` | Output path. Supports `Expand-FileName` tokens. |
| `AtlassianDocumentFileType` | `Full` | `Full` = one file; `Object` = one file per policy plus a TOC index file. |
| `AtlassianTitleProperty` | `Intune documentation` | `<h1>` of the index page. |
| `AtlassianOpenFile` | `$true` | Launch the file with the OS default handler. Set `$false` for CI. |

## Heading anchors (a consumer contract)

Every heading is emitted with an inline anchor macro:

```xml
<h4 id='section-42'><ac:structured-macro ac:name='anchor' ac:schema-version='1'><ac:parameter ac:name=''>section-42</ac:parameter></ac:structured-macro>Contoso Reader</h4>
```

The anchor name has appeared as literal `<!--#section-N-->` text downstream of the
generated file. The source of that rewrite is still unproven: the publishing runbook
parses the body with mshtml before POSTing, and Confluence also transforms storage
format. The inline placement remains the known baseline until those paths are measured
separately.

Three properties are guaranteed, and downstream tooling depends on them:

1. **The `anchor` macro carries the link target.** Confluence discards author-specified
   `id=` attributes when it converts storage format to ADF, so `#section-42` can only
   resolve to an anchor macro or to Confluence's own heading-text-derived anchor.
2. **`id=` always equals the macro name, byte for byte.** Confluence strips it, but
   automation that parses the *generated file* before publishing reads the anchor from
   it - for example to walk an assignment table cell back to its policy heading and
   build a deep link into the published page.
3. **The table of contents links these names, never the heading text.** Anchor names
   contain no character that came from tenant data.

**Naming scheme (breaking-change surface):** `section-N` for every heading, numbered
in emission order over the whole run. `section-N` is unique across a run even in
`Object` mode, because the counter is reset only in
`Invoke-AtlassianPreProcessItems`. Numbering is *not contiguous within the TOC*: a
heading consumes a number whenever it is written to the document, including the
level-6 table captions that the TOC's level cap (4) filters out and the `-SkipTOC`
script captions it never lists. Do not assume contiguity - and do not reconstruct the
numbers by counting TOC entries.

`table-N` is reserved for the `-ToT` caption form of `Add-AtlassianHeader`, which no
call site in this provider uses: table captions here are plain level-6 headings, as
in the HTML provider. Only the Markdown provider passes `-ToT` (and so is the only
output with visible `Table N.` numbering). Enabling it for Atlassian would change
rendered captions, so it belongs with the HTML provider as one formatting decision
rather than a port detail.

The scheme is positional, so inserting one policy shifts every later anchor: links a
user saved from an earlier run of a scheduled export therefore move. A content-derived
scheme (`policy-<objectId>` from the Graph GUID) would be stable and is the natural
follow-up; it needs `Add-AtlassianHeader` to accept an explicit anchor from
`Invoke-AtlassianProcessItem`, with the positional counter kept as the fallback for
headers that have no natural identifier.

**Deployment ordering:** update the module wherever the export runs *before* a consumer
switches to reading `id=` / linking `#section-N`, or its links will point at anchors
that do not exist yet.

## Markup constraints

- **Single-quote every attribute.** Consumers JSON-escape the document body before a
  Confluence client serialises it again; a double-quoted attribute arrives as
  `ac:name=\"anchor\"` and breaks the macro. Every macro in the provider follows this,
  so the body survives a JSON round-trip byte-identically.
- **Storage format is strict XML** and declares only the five XML built-in entities.
  Use numeric references (`&#160;`, never `&nbsp;`), close every macro, and escape
  text that comes from the tenant - `Get-AtlassianXmlText` for headings, TOC labels,
  the title and the document-info lines; `Set-AtlassianText` for table cell values
  (it also wraps XML-looking values in a `code` macro and long text in an `expand`).
  A single bare `&` in a policy name invalidates the whole page body, not one heading,
  and Confluence rejects the upload.
- **In `Object` mode the TOC's `href` carries a file name derived from a policy
  name.** `Get-AtlassianObjectFileName` strips only path-invalid characters, so `&`,
  `'` and `#` survive into it. `Get-AtlassianHref` percent-encodes the file-name
  component (never the `#` that introduces the fragment) and then XML-escapes the
  result; use it rather than interpolating a file name into an attribute.

## Change history

Unreleased (part of 4.0, no shipped version has the earlier behaviour):

- Table of contents entries link the anchor macro each heading emits. They used to
  target a fragment re-derived from the heading text, so duplicate policy names all
  jumped to the first occurrence and characters other than a plain space (`.`, `(`,
  `)`, `:`, `&`, `+`, `,`, U+00A0) leaked into the href unencoded.
- Headings carry an inline `anchor` macro; previously they carried only an `id=` that
  Confluence discards, which nothing could link to. The *shape* of an anchor name is
  unchanged (`section-N`), so a consumer reading `id=` needs no update. A provisional
  preceding-paragraph placement was reverted before release because its downstream
  behavior and blank-line cost had not been measured.
- A heading kept out of the TOC now consumes an anchor number. It used to reuse the
  next listed heading's number, so a document containing script captions (Detection
  script, Requirement scripts) emitted two `section-N` anchors with one name and the
  TOC entry landed on the caption. This shifts the numbers in such documents: another
  reason to read `id=` rather than compute `section-N` from a position.
- Object-mode `href` file names are percent-encoded. A policy name containing `&`
  or `'` used to produce a malformed body, and one containing `#` a link to the wrong
  fragment.
- In the `-ToT` caption path (present but unused, see above) the `Table N. ` prefix
  moved from the id to the visible text, where the Markdown provider puts it.
- Heading text, TOC labels, the title and the document-info lines are XML-escaped.
- Confluence keeps generating its own text-derived anchors, so externally saved
  `#Policy-Name` links still resolve.

The three guarantees above are covered by the project's own test suite.
