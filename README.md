# AutoSort AI for Google Drive

AutoSort AI for Google Drive is an open-source software tool designed to automatically classify academic and educational documents stored in Google Drive using a hybrid decision strategy based on: (i) keyword matching, (ii) semantic similarity with sentence embeddings, and (iii) large language models (LLMs).

The software retrieves files from a source folder, extracts textual content from multiple file formats, applies the three classification strategies, determines a final category using predefined decision rules, and moves each file to its corresponding category folder in Google Drive. Additionally, all classification results are logged into a Google Sheets report for traceability and posterior analysis.

## Software metadata

- **Current version:** 1.0.0  
- **Permanent link to code repository:** https://github.com/miguelangelromerooc-commits/Autosort  
- **Legal code license:** MIT  
- **Programming language:** Python  
- **Operating system:** Windows, Linux, macOS  

**Dependencies:**
- google-api-python-client  
- google-auth-oauthlib  
- nltk  
- sentence-transformers  
- scikit-learn  
- PyMuPDF (fitz)  
- python-docx  
- openpyxl  
- openai / google-generativeai  

**External services:**  
Google Drive API, Google Sheets API, LLM provider API (e.g., OpenAI, Google Gemini)

## Motivation and significance

The organization and retrieval of educational resources stored in cloud-based repositories remains a challenge for instructors and academic institutions. AutoSort AI for Google Drive addresses this problem by providing an automated, explainable and extensible document classification workflow that combines symbolic (keywords) and semantic (embeddings and LLMs) approaches.

This hybrid design improves robustness when dealing with heterogeneous educational materials and supports knowledge management practices in educational contexts, enabling instructors to organize large collections of digital resources with minimal manual intervention.

## Features

- Automatic retrieval of documents from Google Drive  
- Text extraction from PDF, DOCX and XLSX files  
- Hybrid classification strategy:
  - Keyword-based scoring  
  - Embedding-based semantic similarity  
  - LLM-based content interpretation  
- Explainable decision rules for final category assignment  
- Automatic creation and management of category folders  
- Logging of classification results in Google Sheets  
- Modular and extensible architecture  

## Installation

Clone the repository:

```bash
git clone https://github.com/miguelangelromerooc-commits/Autosort.git
cd Autosort
```
Create and activate a virtual environment:
```bash
python -m venv venv
venv\Scripts\activate  
```
Install dependencies:
```bash
pip install -r requirements.txt
```
## Configuration
Google APIs (Drive and Sheets)
- Create a project in Google Cloud Console.
- Enable:
    - Google Drive API
    - Google Sheets API
- Create OAuth 2.0 credentials:
Download the credentials file (from Google Cloud) and name it: "credentials.json".
Note: this file is not included in this repository for security reasons. Each user must create their own OAuth 2.0 credentials in Google Cloud Console and place the downloaded file (renamed as credentials.json) in the project root directory before running the software. 
```
``` 
## Usage
Run the main script:
``` bash
python classify-code.py
```
The software will:

- Authenticate with Google Drive
- Retrieve all files from the source folder
- Extract text from each document
- Apply keyword-based, embedding-based and LLM-based classification
- Determine a final category using predefined decision rules
- Move each file to its corresponding category folder
- Log results into Google Sheets

## Reproducibility and extensibility

- Keyword dictionaries and category definitions can be adapted to other educational domains.
- Similarity thresholds and confidence parameters can be tuned to fit different datasets.
- The LLM module can be replaced with alternative providers or models.
- The modular design allows easy integration of additional classifiers or preprocessing pipelines.

## Limitations
- Classification performance depends on the quality of text extracted from source documents.
- LLM-based classification requires external API access and may incur usage costs.
- Google API quotas may limit large-scale batch processing.

## License
MIT License. See LICENSE for details.
