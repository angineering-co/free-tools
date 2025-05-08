# Message Generator (訊息產生器)

A Google Apps Script add-on for Google Sheets that helps generate messages based on templates and variables.

## Features

- Create custom message templates with variable placeholders
- Define variables and their values in a spreadsheet
- Automatically generate messages by replacing variables in templates
- Simple menu interface in Google Sheets

## Setup

1. Open your Google Sheet
2. Go to Extensions > Apps Script
3. Copy the contents of `Code.ts` into the script editor
4. Save and refresh your spreadsheet
5. You should see a new menu item "訊息產生器" (Message Generator)

## Usage

1. Create two sheets in your spreadsheet:
   - "訊息產生器" (Message Generator): For variables and message generation
   - "情境模板" (Template): For storing message templates

2. In the "訊息產生器" sheet:
   - Columns A-B: Define your variables and their values
   - Columns D-E: Define message keys and their generated content

3. In the "情境模板" sheet:
   - Columns A-B: Define template keys and their content
   - Use `{variableName}` syntax for variables in templates

4. Click "訊息產生器" > "生成訊息" to generate messages

## Template Syntax

Use `{variableName}` in your templates to insert variable values. For example:
