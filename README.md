# PDF to MS Word Replication using Python
## Overview

### This project recreates a given legal PDF document (Mediation Application Form – Form ‘A’) into an MS Word (.docx) file using Python.
### The goal is to replicate the content, layout, structure, spacing, and alignment of the PDF as closely as possible through programmatic document generation.

## Problem Statement

Convert the provided PDF file into an MS Word document by:

1. Preserving the exact structure and order of sections

2. Maintaining headings, numbering, and placeholders

3. Recreating the form layout using tables

4. Avoiding manual editing or auto-conversion tools

## Approach

1. Analyzed the PDF to understand its table-based layout

2. Identified that the form is structured as a single multi-row table

3. Used python-docx to generate the Word document from scratch

4. Recreated:

    Centered headers

   Section titles

   Applicant and Opposite Party details

   Conditional placeholders

   “Details of Dispute” section

    Used table borders and merged cells to match the form layout

## Technologies Used

1. Python 3.x

2. python-docx

3. Google Colab

## Project Structure
pdf-to-docx-replica/
│
├── pdf_to_docx.ipynb
├── Mediation_Application_Form_FINAL.docx
├── django_assignment.pdf
├── README.md

## How to Run

1. Open the Google Colab notebook (pdf_to_docx.ipynb)

2. Upload django_assignment.pdf

3. Install dependencies:

4. Run all cells

5. Download the generated Word file

## Output

1. Mediation_Application_Form_FINAL.docx

2. Fully structured legal form

3. Matches the original PDF layout and content

4. Generated entirely using Python
