# EIFS Voice Data Grid Tool

Voice-first tool for filling the Data grid template.docx document without typing.

## What it does

- Captures spoken commands using browser speech recognition.
- Fills date, photograph number, client, and elevation.
- Fills grid rows with grid letter, grid number, probe, substrate, and comments.
- Adds rows beyond template capacity by cloning the last row format.
- Exports a new .docx file based on Data grid template.docx.

## Hard data rules built in

- Grid letter: A through R, or range format like A-C.
- Grid number: 1 through 14, or range format like 1-14.
- Probe: integer 10 through 40.
- Substrate: F, M, or S.

## Run locally

1. Install dependencies:

	pip install -r requirements.txt

2. Start server:

	uvicorn app:app --reload

3. Open:

	http://127.0.0.1:8000

## Voice command examples

- document name Queens Landing B24 Grid
- date April 27 2026
- photograph 14
- client Queens Landing HOA
- elevation South
- grid letter A-C
- grid number 1-14
- probe 22
- substrate F
- comment minor moisture at seam
- add row
- next row
- previous row
- go back
- go forward
- delete word
- clear field
- save document

## Render deploy

render.yaml is configured for a Python web service.

