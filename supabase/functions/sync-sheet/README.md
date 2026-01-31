# Google Sheets Sync (Supabase Edge Function)

This function updates a Google Spreadsheet so **each order has its own sheet tab** that contains
all purchase details, including links and image URLs.

## Required environment variables

- `SUPABASE_URL`
- `SUPABASE_SERVICE_ROLE_KEY`
- `GOOGLE_SHEET_ID` (your spreadsheet ID)
- `GOOGLE_SERVICE_ACCOUNT_EMAIL`
- `GOOGLE_PRIVATE_KEY` (service account private key, with `\n` escaped)

## Spreadsheet

Spreadsheet ID used in production:

```
1d2yQ_ovd7w0CamppPpnvoQ2fb_3U6Q-m19tv7As1jiQ
```

## What it writes

Each sheet (tab) is named after the **order name** (or order ID if name is missing), and includes:

- order name
- purchase fields (customer, qty, price, pickup point, note)
- pickup + collection flags and timestamps
- purchase links (newline-separated)
- image URLs (newline-separated)

## How to trigger

Create Supabase **Database Webhooks** for the following tables:

- `orders`
- `purchases`
- `purchase_links`
- `purchase_images`

Each webhook should call the Edge Function and pass an `order_id` in the payload.

Example payload shape (any of these will work):

```json
{ "order_id": "<uuid>" }
```

or

```json
{ "record": { "order_id": "<uuid>" } }
```
