# app-script

This repository houses various Google Apps Script projects developed for different purposes.

## Overview

This repository serves as a central location to organize and manage my collection of Google Apps Script projects. Each project within this repository is intended to automate tasks, extend the functionality of Google Workspace applications (like Google Sheets, Docs, Forms, etc.), or integrate with other services.

## Contents

Each sub-directory within this repository will typically represent a distinct Google Apps Script project. You might find folders for:

* **gsheet-automation:** Scripts specifically designed to automate tasks within Google Sheets (e.g., data processing, reporting, email sending).
* **gdrive-management:** Scripts for managing files and folders within Google Drive.
* **gmail-utilities:** Scripts to enhance Gmail functionality (e.g., automated email responses, label management).
* **calendar-integrations:** Scripts that interact with Google Calendar.

Within each project directory, you will generally find:

* `.gs` files: The Google Apps Script code files.
* `appsscript.json`: The project manifest file, defining project settings, scopes, and triggers.
* Potentially other supporting files (e.g., configuration files, READMEs specific to the project).

## Usage

This repository is primarily for personal organization and version control of my Google Apps Script projects. If you are viewing this repository, it's likely either because I've shared it or it's public. Feel free to browse the code, but note that these scripts are tailored to my specific needs and may require modification to be useful in other contexts.

## Development

These projects are developed locally using VS Code and managed with CLASP (Command Line Apps Script Projects). This setup allows for a more robust development environment with features like Git version control.

To work with these projects locally (if you have access and CLASP configured):

1.  **Clone this repository:** `git clone https://github.com/emirevec/app-script.git`
2.  **Navigate to a specific project directory:** `cd app-script/gsheet-automation`
3.  **If the project was initialized with CLASP, you might find a `.clasp.json` file.** This links the local directory to the remote Google Apps Script project.
4.  **You can then use CLASP commands** (e.g., `clasp push`, `clasp pull`, `clasp open`) to manage the Apps Script project.

## Contributing

As this is primarily a personal repository, external contributions are not generally expected. However, if you have suggestions or find issues, feel free to open an issue to discuss them.
