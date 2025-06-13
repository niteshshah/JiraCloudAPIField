# Jira Cloud Excel SharePoint Field

This repository contains a basic Atlassian Connect app that exposes a Jira custom field. The field loads selectable options from an Excel sheet stored in SharePoint via the Microsoft Graph API.

## Setup

1. Install dependencies:
   ```bash
   npm install
   ```
2. Copy `.env.example` to `.env` and fill in your Microsoft Azure and SharePoint details.
3. Start the app:
   ```bash
   node server.js
   ```
4. Install the app in Jira Cloud using the URL to `atlassian-connect.json`.

The `/excel-options` endpoint reads the Excel sheet and returns available values for the custom field.
