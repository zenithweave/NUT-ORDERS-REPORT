# Custom Shopify Order Export

A Node.js application that provides custom order export functionality for Shopify stores, matching WooCommerce CSV format.

## Features

- Export Shopify orders in WooCommerce CSV format
- Filter by date range and order status
- Complete order and customer data
- Compatible with existing workflows

## Setup

1. Install dependencies:
```bash
npm install
```

2. Configure environment variables in `.env`:
```
SHOPIFY_SHOP_NAME=your-store.myshopify.com
SHOPIFY_ADMIN_ACCESS_TOKEN=shpat_your_token_here
PORT=3000
```

3. Start the server:
```bash
npm start
```

## API Endpoints

- `POST /export-orders` - Export orders with filters
- `GET /order-statuses` - Get available order statuses
- `GET /health` - Health check

## Deployment

This app is configured for Railway deployment. Push to GitHub and connect to Railway.
