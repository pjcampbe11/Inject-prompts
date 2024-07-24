# Inject Hidden Prompt Script

## Overview
This script allows you to inject hidden prompts into Microsoft 365 documents (Word, Excel, PowerPoint, and OneNote). The payload can be specified as a text string or as a file path.

## Requirements
- Python 3.x
- `zipfile` and `os` modules (standard library)

## Usage
```bash
python inject_hidden_prompt.py <doc_type> <location> <payload> [--file]
```

### Arguments
- `doc_type`: The type of document. Choices are `word`, `excel`, `powerpoint`, or `onenote`.
- `location`: The location within the document to inject the payload. Choices vary by document type.
- `payload`: The payload to be injected. This can be a text string or a file path.
- `--file`: Optional flag to indicate if the payload is a file path.

### Example Commands

#### Word Document
Injecting a hidden prompt into the `_rels` section of a Word document:

```bash
python inject_hidden_prompt.py word rels "hidden_prompt_text"
```

Injecting a hidden prompt from a file into the `document` section of a Word document:

```bash
python inject_hidden_prompt.py word document path/to/payload.txt --file
```

#### Excel Document
Injecting a hidden prompt into the `workbook` section of an Excel document:

```bash
python inject_hidden_prompt.py excel workbook "hidden_prompt_text"
```

Injecting a hidden prompt from a file into the `styles` section of an Excel document:

```bash
python inject_hidden_prompt.py excel styles path/to/payload.txt --file
```

#### PowerPoint Document
Injecting a hidden prompt into the `presentation` section of a PowerPoint document:

```bash
python inject_hidden_prompt.py powerpoint presentation "hidden_prompt_text"
```

Injecting a hidden prompt from a file into the `slide1` section of a PowerPoint document:

```bash
python inject_hidden_prompt.py powerpoint slide1 path/to/payload.txt --file
```

#### OneNote Document
Injecting a hidden prompt into the `onetoc2` section of a OneNote document:

```bash
python inject_hidden_prompt.py onenote onetoc2 "hidden_prompt_text"
```

Injecting a hidden prompt from a file into the `section` section of a OneNote document:

```bash
python inject_hidden_prompt.py onenote section path/to/payload.txt --file
```

### Document Types and Locations

#### Word (.docx)
- `rels`: `_rels/.rels`
- `docProps`: `docProps/core.xml`
- `document`: `word/document.xml`
- `fontTable`: `word/fontTable.xml`
- `settings`: `word/settings.xml`
- `styles`: `word/styles.xml`
- `theme`: `word/theme/theme1.xml`
- `webSettings`: `word/webSettings.xml`
- `docRels`: `word/_rels/document.xml.rels`
- `contentTypes`: `[Content_Types].xml`

#### Excel (.xlsx)
- `rels`: `_rels/.rels`
- `docProps`: `docProps/core.xml`
- `workbook`: `xl/workbook.xml`
- `styles`: `xl/styles.xml`
- `sharedStrings`: `xl/sharedStrings.xml`
- `theme`: `xl/theme/theme1.xml`
- `sheet1`: `xl/worksheets/sheet1.xml`
- `workbookRels`: `xl/_rels/workbook.xml.rels`
- `contentTypes`: `[Content_Types].xml`

#### PowerPoint (.pptx)
- `rels`: `_rels/.rels`
- `docProps`: `docProps/core.xml`
- `presentation`: `ppt/presentation.xml`
- `slide1`: `ppt/slides/slide1.xml`
- `slideLayouts`: `ppt/slideLayouts/slideLayout1.xml`
- `slideMasters`: `ppt/slideMasters/slideMaster1.xml`
- `notesSlide`: `ppt/notesSlides/notesSlide1.xml`
- `theme`: `ppt/theme/theme1.xml`
- `presentationRels`: `ppt/_rels/presentation.xml.rels`
- `contentTypes`: `[Content_Types].xml`

#### OneNote (.one)
- `onetoc2`: `openNotebookStructure.onetoc2`
- `section`: `sectionFolder`
