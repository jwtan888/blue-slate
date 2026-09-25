# Trip Claims

An offline-first mobile web app for capturing business-trip expenses travel reimbursement workbook.

## Run locally

```bash
npm install
npm run dev -- --host 0.0.0.0
```

Open the displayed URL on the iPhone while it is on the same Wi-Fi, or deploy the `dist/` output to an HTTPS host. In Safari, use **Share → Add to Home Screen**.

## Workflow

1. Enter the trip details.
2. Add each expense with its date, currency, amount, description, payee, and optional receipt photo.
3. Enter cash advances and returned VND notes when applicable.
4. Download the Excel claim and the receipt ZIP after the trip.

Trip data is stored locally on the device. The supplied workbook is kept in `public/template.xlsx` and is used as the export base.

The VND → MYR default is the rate in the supplied workbook (`0.00015`). USD stays unavailable until an approved USD → MYR rate is entered.

## Verification

```bash
npm run test:smoke
```

The smoke test renders the app in a DOM harness, saves a sample expense, exports the workbook, and checks the exported sheet, date, formulas, advance, cash return, and stale-template cleanup.
