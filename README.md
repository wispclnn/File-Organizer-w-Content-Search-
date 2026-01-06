
An intelligent, content-aware file management application designed to help users organize, search, and manage large collections of digital files efficiently.

## Overview

Traditional file managers rely heavily on filenames, making it difficult to locate files when names are unclear or inconsistent. This project addresses that limitation by introducing **content-aware search**, **rule-based automation**, and **file preview** capabilities.

The system analyzes the actual content of files (PDFs, documents, images via OCR, etc.) and allows users to retrieve files based on what is *inside* them, not just their names.

## Key Features

* **Content-Aware Search**

  * Searches files using extracted text, not just filenames
  * Uses **TF-IDF** and **Cosine Similarity** for relevance ranking

* **OCR Support**

  * Extracts text from images using Tesseract OCR

* **Rule-Based Automation**

  * Automatically move, copy, or delete files based on user-defined rules

* **File Preview**

  * Preview supported file types without opening external applications

* **Tagging System**

  * Assign custom tags to files for easier organization

* **Undo Functionality**

  * Revert recent file operations (batch-based undo)

* **Folder Shortcuts Panel**

  * Quick access to frequently used folders

##How It Works

1. The system scans files within a selected folder.
2. Text is extracted from supported file types (PDF, DOCX, TXT, images via OCR).
3. Extracted content is cached locally for faster future searches.
4. User search queries are converted into TF-IDF vectors.
5. Cosine similarity is computed between the query and document vectors.
6. Files are ranked and displayed based on relevance.

## Tech Used

* **Python**
* **Tkinter** (GUI)
* **scikit-learn** (TF-IDF, cosine similarity)
* **pytesseract** (OCR)
* **PyPDF2 / python-docx** (content extraction)

## Limitations

* Undo history resets when the application is closed
* OCR accuracy depends on image quality
* Content extraction is limited to supported file formats

## Intended Use

This application is suitable for students, professionals, and users who manage large numbers of documents and need faster, more reliable file retrieval.


