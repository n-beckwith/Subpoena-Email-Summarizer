# Subpoena Email Summarizer

## Overview
The Email Summarizer is a Python desktop application built with Tkinter. It scans designated file system directories recursively for Microsoft Outlook message files (.msg), parses metadata and body content, generates concise AI summaries via the GitHub Models API endpoint, and renders live processing metrics.
Extracted records can be exported directly into structured Microsoft Excel Workbooks (.xlsx) or copied to the clipboard.

## Key Architectural Features
*	Cloud API Integration: Utilizes the OpenAI Python SDK routed through GitHub Models (openai/gpt-4o-mini). Includes direct exception handling for HTTP 429 Rate Limit errors and token quota limits.
*	Tokens: Automatically skips API summary generation for short emails (fewer than 300 characters or ~40 words), calendar invitations, auto-replies, and quick-replies.
*	Real-Time State: Saves execution state (source folder path, next resume index, execution status) to pipeline_state.json. If a rate limit is triggered, the engine calculates the exact failure index, updates UI automatically, and recalculates the remaining limits.
*	Threaded UI: The main extraction loop executes on a background thread, keeping the GUI responsive during long-running operations.
*	Live Feed Table: Displays the metadata in a custom Treeview component featuring row checkboxes, toggle/invert batch selection options, and full-body text preview windows on double-click.
 
## System Requirements and Prerequisites
### User Environment
*	Operating System: Windows 10 / Windows 11 (64-bit)
*	Network Requirements: Active Internet connection (HTTPS port 443 outbound access to https://models.github.ai/inference)
*	API Key: A valid GitHub Personal Access Token configured inside the source file or environment variables. (The current token will expire on AUGUST 21, 2026)

### Developer Dependencies
To run or recompile the raw source code, install the required packages using Python 3.10 or higher: 
pip install openpyxl pandas extract_msg beautifulsoup4 openai pyinstaller

## User Interface and Operational Guide
### Configuration Panel
1. Source Folder: Click Browse... to pick the folder containing your .msg files. The path is saved immediately to disk and auto-loads on future launches.
2. Start Index: Designates the target start item index. Defaults to 1. If an execution stops early due to an API limit, this field is changed to the next unprocessed item index.
3. Limit Total Items: Optional field to restrict execution length to a specific number of files. On early rate-limit termination, this field automatically updates to reflect the remaining unprocessed balance.

### Control and Action Buttons
*	START: Initiates directory scanning and pipeline execution.
*	Select and Clear All Rows Selections: Selects or deselects all rows currently loaded in the feed table.
*	Invert Selection: Flips current row check states (checked items become unchecked, unchecked items become checked).
*	EXPORT LOG DATA: Drops down options for Excel document generation, clipboard copying, and row filtering.

### Table Operations
*	Single Click Column 1: Toggles row checkmark state.
*	Double-Click Row Item: Opens a detached viewer window displaying the full, raw body text of the selected message.

## Data Export Specifications
### Excel Export Format (.xlsx)
When exporting to Excel through EXPORT LOG DATA --> Excel Document Exports:
*	Headers: FileName, Subject, Sender, Recipient, CC, Attachments, Date, Time, Summary.
*	Formatting: Styled headers (#1F4E78 fill, white bold font), thin borders, automatic text wrapping, and explicit column widths.
*	Deduplication Logic: Re-running an export over an existing spreadsheet updates the matching file entries without creating duplicate rows.
*	Rate-Limit Indicator: Appends a highlighted notice banner at the bottom of the worksheet if execution halted before completion due to API rate limits.

### Clipboard Copy Operations
Accessible under EXPORT LOG DATA --> Clipboard Operations:
*	Exports checked rows (or all rows if none are checked) in tab-delimited text format, suitable for pasting directly into text editors, Microsoft Word, or database fields.
*	Includes an option to toggle field headers on or off.

| Error Handling and Troubleshooting | Cause | Solution |
| --- | --- | --- |
| **Configuration Error: Paste GitHub Token** | `GITHUB_TOKEN` variable is empty or unconfigured. | Open source code and assign a valid token string to `GITHUB_TOKEN`. |
| **Write Lock Error** | The target `.xlsx` file is currently open in Microsoft Excel or another program. | Close the target Excel document and retry. |
| **API Limit Reached** | Exhausted GitHub Models API request/token rate limits (HTTP 429). | Wait for the indicated reset window. The application automatically saves your resume Start Index and Limit values. Click START once the quota resets. |
| **NameError on Startup** | Widget instantiation sequence mismatch in source code. | Ensure the state logic (`load_pipeline_state`) is placed below widget definitions (`ent_folder`, `ent_start_idx`). |

## Recompilation and Build Procedure
To recompile updated Python source code into a standalone executable package:
1.	Open Command Prompt in the script folder.
2.	Test the script execution directly:
python Email_Extractor.py
3.	Compile into a windowed executable: 
pyinstaller --noconfirm --onedir --windowed --name "Email_Extractor" Email_Extractor.py
4.	Access the generated executable package in the dist/Email_Extractor directory.
