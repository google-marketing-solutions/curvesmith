# CurveSmith

**This is not an official Google product.**

## Overview

This Apps Script solution streamlines the bulk creation and upload of custom
delivery curves to Google Ad Manager (GAM) line items.

The solution leverages your existing Google Sheets credentials to access GAM
networks for which you already have authorization.

Built with: [Google Apps Script](https://www.google.com/script/start/),
[TypeScript](https://www.typescriptlang.org/)

### Prerequisites:

*   Access to Google Sheets
*   Access to one or more GAM networks
*   Basic understanding of how
    [custom delivery curves](https://support.google.com/admanager/answer/9293326?hl=en)
    work in GAM

## Deploy

To get your own copy of this solution, you can make a copy of this
[Google Sheet](https://docs.google.com/spreadsheets/d/1caV1a76I9Pel_TR_iwxUz4a0K9LCKK3L0MDOGwNXGvg/edit?usp=sharing)
or deploy entirely from code as detailed below.

## Deploy manually

After cloning the repository, open a terminal in the repository directory.

Install dependencies with npm:

```sh
$ npm install
```

Build the library:

```sh
$ npm run build
```

Use [clasp](https://developers.google.com/apps-script/guides/clasp#installation)
from the top level of the repository to create a Google Sheets script:

```sh
$ clasp login
```

This may open a browser window and ask that you authenticate using your Google
account credentials.

```sh
$ clasp create
```

Choose "sheets" to create a Google Sheets container and an associated client
Apps Script. This may fail if you are a first time user, so follow instructions
for [enabling the Apps Script API](https://script.google.com/home/usersettings)
if any are provided. Take note of the URLs provided by clasp as the first link
is to the spreadsheet that will house the solution script.

Once clasp is configured, build and deploy with:

```sh
$ npm run deploy
```

## Configure the spreadsheet

You will need to initialize the spreadsheet created earlier:

*   Open the Google Sheets URL obtained from clasp and within the spreadsheet
    choose an existing sheet or create a new one to use for basic configuration
    properties.

    *   Create a one-cell
        [named range](https://support.google.com/docs/answer/63175) called
        `NETWORK_CODE` and input a GAM network code that you are authorized to
        access.
    *   Create a one-cell named range called `API_VERSION` and input the Ad
        Manager API version that should be used (eg. v0000 - use the latest from
        [here](https://developers.google.com/ad-manager/api/rel_notes)).
    *   Create a one-cell named range called `TEMPLATE_SHEET_NAME` and input the
        name of a sheet that will be configured in the next section. This
        template will be used to define multiple schedules for different line
        item batches and allow for custom formatting.

*   Create a new sheet with the template name you defined in the previous range.

In your template sheet, define the following named ranges:

*   `AD_UNIT_ID` (1 cell)
*   `GOAL_TYPE` (1 cell)
*   `SCHEDULED_EVENTS` (4 columns, 50+ rows, or as many as you anticipate
    needing)
*   `LINE_ITEMS` (6 columns, 50+ rows, or as many as you anticipate needing)
*   `SELECT_ALL` (1 cell, insert a checkbox (`Insert -> Checkbox`) for easy
    toggling)

## Usage

If the project has been deployed and configured correctly, a new menu item
called `Custom Curves` will appear a few seconds after opening the container
spreadsheet. If it doesn't appear, you can try refreshing the page.

To use the solution for the currently configured GAM network, select `Custom
Curves > Show Sidebar`.

## Disclaimer

This is not an officially supported Google product. The code shared here is not
formally supported by Google and is provided only as a reference.

This solution allows users with Ad Manager network access to interact with their
Ad Manager data through Google Sheets. However, it is important to understand
that certain line item information will be visible within the Google Sheets
spreadsheet to anyone who has access to it, regardless of whether they have
permission to view that data in Ad Manager itself.

The specific line item details that may be visible include:

-   Line item ID
-   Line item name
-   Flight dates (start and end dates)
-   Impression goal

To maintain data confidentiality, we strongly recommend that users avoid
including any sensitive or confidential terms when naming their line items in Ad
Manager.

By using this solution, you acknowledge and accept the potential visibility of
this basic line item data within the Google Sheets environment.
