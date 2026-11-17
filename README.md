# Import Sheet Checker

Checks shipper invoice (Excel), Bill of Entry (PDF) and CHA tax invoice + debit note (PDF) against the import sheet format.

Everything runs **in the browser**. No login, no API key, no server, files are never uploaded.

## Run locally
```
npm install
npm run dev      # http://localhost:3000
```
## Deploy
Push to GitHub and import into Vercel (no environment variables needed).

## How it reads documents
Fixed labels and positions, not AI:
- Shipper invoice `.xls/.xlsx`: Invoice Number, Date, Shipper/Exporter, QTY / TOTAL PRICE / G.W(KG) item columns.
- BOE PDF (CBE-XIII): OOC date, CBE-XIII number, HAWB, consignor, rate of exchange, duty, IGST (sum of items), invoice number, airline.
- CHA PDF: invoice / debit note no. and date, HAWB list, flight, chargeable weight, freight rate x weight, charge lines, GST, TOTAL.

Anything it cannot find is shown as MISSING and is never guessed. Scanned (image) PDFs have no text and cannot be read.

## Excel format
The download is built from `public/template.xlsx` (your import sheet: header rows, colours, widths, merged CHA groups, "Format" row).
To change the layout, replace that file. Keep the header in rows 1-2, one sample data row at row 3 (its styles are copied) and the row whose column A says `Format`.
