const express = require('express');
const cors = require('cors');
const { createObjectCsvWriter } = require('csv-writer');
const moment = require('moment');
const path = require('path');
require('dotenv').config();

const app = express();
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// Shopify API configuration
const SHOPIFY_SHOP_NAME = process.env.SHOPIFY_SHOP_NAME;
const SHOPIFY_ADMIN_ACCESS_TOKEN = process.env.SHOPIFY_ADMIN_ACCESS_TOKEN;

// Serve frontend
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Health check endpoint
app.get('/health', (req, res) => {
  res.json({ status: 'OK', timestamp: new Date().toISOString() });
});

// Helper function to make Shopify API requests
async function fetchShopifyData(endpoint) {
  const url = `https://${SHOPIFY_SHOP_NAME}/admin/api/2023-10/${endpoint}`;
  const response = await fetch(url, {
    headers: {
      'X-Shopify-Access-Token': SHOPIFY_ADMIN_ACCESS_TOKEN,
      'Content-Type': 'application/json'
    }
  });
  
  if (!response.ok) {
    throw new Error(`Shopify API error: ${response.status} ${response.statusText}`);
  }
  
  return response.json();
}

// Transform Shopify order data to WooCommerce CSV format
function transformOrderToWooFormat(order, lineItem) {
  return {
    "Order Number": order.name || order.id,
    "Order Status": order.financial_status || 'pending',
    "Order Date": moment(order.created_at).format('YYYY-MM-DD HH:mm'),
    "Customer Note": order.note || '',
    "First Name (Billing)": order.billing_address?.first_name || '',
    "Last Name (Billing)": order.billing_address?.last_name || '',
    "Company (Billing)": order.billing_address?.company || '',
    "Address 1&2 (Billing)": `${order.billing_address?.address1 || ''} ${order.billing_address?.address2 || ''}`.trim(),
    "City (Billing)": order.billing_address?.city || '',
    "State Code (Billing)": order.billing_address?.province_code || '',
    "Postcode (Billing)": order.billing_address?.zip || '',
    "Country Code (Billing)": order.billing_address?.country_code || '',
    "Email (Shipping)": order.shipping_address?.email || order.email || '',
    "Paid Date": order.processed_at ? moment(order.processed_at).format('YYYY-MM-DD HH:mm') : '',
    "Phone (Shipping)": order.shipping_address?.phone || order.phone || '',
    "Phone (Billing)": order.billing_address?.phone || order.phone || '',
    "First Name (Shipping)": order.shipping_address?.first_name || order.billing_address?.first_name || '',
    "Last Name (Shipping)": order.shipping_address?.last_name || order.billing_address?.last_name || '',
    "Address 1&2 (Shipping)": `${order.shipping_address?.address1 || order.billing_address?.address1 || ''} ${order.shipping_address?.address2 || order.billing_address?.address2 || ''}`.trim(),
    "City (Shipping)": order.shipping_address?.city || order.billing_address?.city || '',
    "State Code (Shipping)": order.shipping_address?.province_code || order.billing_address?.province_code || '',
    "Postcode (Shipping)": order.shipping_address?.zip || order.billing_address?.zip || '',
    "Country Code (Shipping)": order.shipping_address?.country_code || order.billing_address?.country_code || '',
    "Payment Method Title": order.payment_gateway_names?.[0] || 'Unknown',
    "Cart Discount Amount": order.total_discounts || '0',
    "Order Subtotal Amount": order.subtotal_price || '0',
    "Shipping Method Title": order.shipping_lines?.[0]?.title || 'Standard',
    "Order Shipping Amount": order.total_shipping_price_set?.shop_money?.amount || order.total_shipping_price || '0',
    "Order Refund Amount": '0',
    "Order Total Amount": order.total_price || '0',
    "Order Total Tax Amount": order.total_tax || '0',
    "SKU": lineItem.sku || '',
    "Item #": lineItem.id,
    "Item Name": lineItem.name || lineItem.title,
    "Quantity (- Refund)": lineItem.quantity || 1,
    "Item Cost": lineItem.price || '0',
    "Coupon Code": order.discount_codes?.map(dc => dc.code).join(', ') || '',
    "Discount Amount": lineItem.total_discount || '0',
    "Discount Amount Tax": '0',
    "City Code": '',
    "Carrier Code": ''
  };
}

// Export orders endpoint
app.post('/export-orders', async (req, res) => {
  try {
    const { startDate, endDate, status } = req.body;
    
    let queryParams = 'limit=250';
    if (startDate) queryParams += `&created_at_min=${startDate}`;
    if (endDate) queryParams += `&created_at_max=${endDate}`;
    if (status) queryParams += `&status=${status}`;
    
    const ordersData = await fetchShopifyData(`orders.json?${queryParams}`);
    const orders = ordersData.orders;
    
    if (!orders || orders.length === 0) {
      return res.status(404).json({ error: 'No orders found' });
    }
    
    const csvData = [];
    for (const order of orders) {
      for (const lineItem of order.line_items) {
        csvData.push(transformOrderToWooFormat(order, lineItem));
      }
    }
    
    const csvWriter = createObjectCsvWriter({
      path: 'orders_export.csv',
      header: [
        {id: 'Order Number', title: 'Order Number'},
        {id: 'Order Status', title: 'Order Status'},
        {id: 'Order Date', title: 'Order Date'},
        {id: 'Customer Note', title: 'Customer Note'},
        {id: 'First Name (Billing)', title: 'First Name (Billing)'},
        {id: 'Last Name (Billing)', title: 'Last Name (Billing)'},
        {id: 'Company (Billing)', title: 'Company (Billing)'},
        {id: 'Address 1&2 (Billing)', title: 'Address 1&2 (Billing)'},
        {id: 'City (Billing)', title: 'City (Billing)'},
        {id: 'State Code (Billing)', title: 'State Code (Billing)'},
        {id: 'Postcode (Billing)', title: 'Postcode (Billing)'},
        {id: 'Country Code (Billing)', title: 'Country Code (Billing)'},
        {id: 'Email (Shipping)', title: 'Email (Shipping)'},
        {id: 'Paid Date', title: 'Paid Date'},
        {id: 'Phone (Shipping)', title: 'Phone (Shipping)'},
        {id: 'Phone (Billing)', title: 'Phone (Billing)'},
        {id: 'First Name (Shipping)', title: 'First Name (Shipping)'},
        {id: 'Last Name (Shipping)', title: 'Last Name (Shipping)'},
        {id: 'Address 1&2 (Shipping)', title: 'Address 1&2 (Shipping)'},
        {id: 'City (Shipping)', title: 'City (Shipping)'},
        {id: 'State Code (Shipping)', title: 'State Code (Shipping)'},
        {id: 'Postcode (Shipping)', title: 'Postcode (Shipping)'},
        {id: 'Country Code (Shipping)', title: 'Country Code (Shipping)'},
        {id: 'Payment Method Title', title: 'Payment Method Title'},
        {id: 'Cart Discount Amount', title: 'Cart Discount Amount'},
        {id: 'Order Subtotal Amount', title: 'Order Subtotal Amount'},
        {id: 'Shipping Method Title', title: 'Shipping Method Title'},
        {id: 'Order Shipping Amount', title: 'Order Shipping Amount'},
        {id: 'Order Refund Amount', title: 'Order Refund Amount'},
        {id: 'Order Total Amount', title: 'Order Total Amount'},
        {id: 'Order Total Tax Amount', title: 'Order Total Tax Amount'},
        {id: 'SKU', title: 'SKU'},
        {id: 'Item #', title: 'Item #'},
        {id: 'Item Name', title: 'Item Name'},
        {id: 'Quantity (- Refund)', title: 'Quantity (- Refund)'},
        {id: 'Item Cost', title: 'Item Cost'},
        {id: 'Coupon Code', title: 'Coupon Code'},
        {id: 'Discount Amount', title: 'Discount Amount'},
        {id: 'Discount Amount Tax', title: 'Discount Amount Tax'},
        {id: 'City Code', title: 'City Code'},
        {id: 'Carrier Code', title: 'Carrier Code'}
      ]
    });
    
    await csvWriter.writeRecords(csvData);
    
    const fs = require('fs');
    const csvContent = fs.readFileSync('orders_export.csv');
    
    res.setHeader('Content-Type', 'text/csv');
    res.setHeader('Content-Disposition', 'attachment; filename="orders_export.csv"');
    res.send(csvContent);
    
  } catch (error) {
    console.error('Export error:', error);
    res.status(500).json({ error: error.message });
  }
});

// Get available order statuses
app.get('/order-statuses', async (req, res) => {
  try {
    const statuses = [
      { value: 'any', label: 'Any Status' },
      { value: 'open', label: 'Open' },
      { value: 'closed', label: 'Closed' },
      { value: 'cancelled', label: 'Cancelled' },
      { value: 'archived', label: 'Archived' }
    ];
    res.json(statuses);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
