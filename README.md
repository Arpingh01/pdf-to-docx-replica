#PDF to MS Word Replication using Python
Overview

This project recreates a given legal PDF document (Mediation Application Form – Form ‘A’) into an MS Word (.docx) file using Python.
The goal is to replicate the content, layout, structure, spacing, and alignment of the PDF as closely as possible through programmatic document generation.

Problem Statement

Convert the provided PDF file into an MS Word document by:

Preserving the exact structure and order of sections

Maintaining headings, numbering, and placeholders

Recreating the form layout using tables

Avoiding manual editing or auto-conversion tools

Approach

Analyzed the PDF to understand its table-based layout

Identified that the form is structured as a single multi-row table

Used python-docx to generate the Word document from scratch

Recreated:

Centered headers

Section titles

Applicant and Opposite Party details

Conditional placeholders

“Details of Dispute” section

Used table borders and merged cells to match the form layout

Technologies Used

Python 3.x

python-docx

Google Colab

Project Structure
pdf-to-docx-replica/
│
├── pdf_to_docx.ipynb
├── Mediation_Application_Form_FINAL.docx
├── django_assignment.pdf
├── README.md

How to Run

Open the Google Colab notebook (pdf_to_docx.ipynb)

Install dependencies:

pip install python-docx


Run all cells

Download the generated Word file

Output

Mediation_Application_Form_FINAL.docx

Fully structured legal form

Matches the original PDF layout and content

Generated entirely using Python
