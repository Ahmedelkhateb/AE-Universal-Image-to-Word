# AE Universal Image to Word

> 🚧 **Work in Progress — Under Active Development**

An experimental Python project that aims to convert document images into editable Microsoft Word documents using OCR and document layout analysis.

## Project Overview

The goal of this project is to reconstruct document content from images into editable Word files, rather than simply inserting the original image into a Word document.

## Current Technologies

- Python 3.11+
- PaddlePaddle
- PaddleOCR
- PP-StructureV3
- OpenCV
- python-docx

## Current Features

- Document image loading and processing.
- Document orientation detection.
- Document layout analysis.
- OCR text extraction.
- Table detection and reconstruction.
- Editable Word text generation.
- Native Word table generation.
- CPU processing support.

## Development Status

<details>
<summary>Current development status</summary>

This project is still under active development and is NOT production-ready.

Areas that require further development and testing include:

- Improving OCR accuracy.
- Improving Arabic and English text recognition.
- Better handling of complex document layouts.
- Improving table reconstruction.
- Preserving text positioning and formatting.
- Testing different image qualities and document types.
- Improving stability and error handling.

The current implementation is experimental. Output quality may vary depending on the input document.

</details>

## Installation

Python 3.11 or newer is recommended.

Install dependencies:

```bash
pip install -r requirements.txt

- Document layout analysis.
- OCR text extraction.
- Table detection and reconstruction.
- Editable Word text generation.
- Native Word table generation.
- CPU processing support.

## Development Status

<details>
<summary>Current development status</summary>

This project is still under active development and is NOT production-ready.

Areas that require further development and testing include:

- Improving OCR accuracy.
- Improving Arabic and English text recognition.
- Better handling of complex document layouts.
- Improving table reconstruction.
- Preserving text positioning and formatting.
- Testing different image qualities and document types.
- Improving stability and error handling.

The current implementation is experimental. Output quality may vary depending on the input document.

</details>

## Installation

Python 3.11 or newer is recommended.

Install dependencies:

```bash
pip install -r requirements.txt

Additional model and runtime requirements may apply depending on the PaddleOCR environment

## Usage

python image_to_editable_word.py input.jpg

Specify an output file:

python image_to_editable_word.py input.jpg -o output.docx

## Project Goal
The long-term goal is to build a general-purpose tool that can reconstruct document images as genuinely editable Word documents while preserving their structure and visual layout as accurately as possible.

## Disclaimer
This repository represents an ongoing development project. Features, accuracy, compatibility, and output quality are subject to change.

Not intended for production use at this stage.

## License

The original code in this repository is licensed under the **MIT License**. See the [LICENSE](LICENSE) file for the full license text.

Copyright (c) 2026 Ahmed Elkhateb.

Third-party packages and their components remain subject to their respective licenses. The MIT license in this repository does not replace or relicense those dependencies.
