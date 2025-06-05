# rcurto_investments

This repository contains the Google Apps Script project files and supporting data.

The Apps Script now computes trade profit/loss values in code rather than relying on formulas in the sheet.
Each add, update, or delete action performed via the web app is recorded in a `TradeInteractions` sheet so you can review a history of changes.

## Running Tests

A small Jest-based test suite validates analytics helper calculations.

Install dependencies (once):

```bash
npm install
```

Execute the tests with:

```bash
npm test
```

