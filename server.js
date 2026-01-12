const express = require('express');
const cors = require('cors');
const axios = require('axios');
const { createObjectCsvWriter } = require('csv-writer');
const moment = require('moment');
const path = require('path');
const ExcelJS = require('exceljs');
require('dotenv').config();

const app = express();
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// Shopify API configuration
const SHOPIFY_SHOP_NAME = process.env.SHOPIFY_SHOP_NAME;
const SHOPIFY_ADMIN_ACCESS_TOKEN = process.env.SHOPIFY_ADMIN_ACCESS_TOKEN;

// Verify required environment variables
if (!SHOPIFY_SHOP_NAME || !SHOPIFY_ADMIN_ACCESS_TOKEN) {
  console.error('❌ Missing required environment variables');
  process.exit(1);
}

// Serve dashboard as main page
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'dashboard.html'));
});

// Keep simple export tool accessible
app.get('/simple', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Health check endpoint
app.get('/health', (req, res) => {
  res.json({ status: 'OK', timestamp: new Date().toISOString() });
});

// Helper function to make Shopify API requests using axios
async function fetchShopifyData(endpoint) {
  try {
    const url = `https://${SHOPIFY_SHOP_NAME}/admin/api/2024-01/${endpoint}`;
    const response = await axios.get(url, {
      headers: {
        'X-Shopify-Access-Token': SHOPIFY_ADMIN_ACCESS_TOKEN,
        'Content-Type': 'application/json'
      }
    });
    return response.data;
  } catch (error) {
    console.error('Shopify API error:', error.response?.data || error.message);
    throw error;
  }
}

// Helper function to fetch ALL orders with pagination
// Note: Shopify API maximum is 250 orders per request
async function fetchAllOrders(queryParams) {
  let allOrders = [];
  let pageInfo = null;
  let hasNextPage = true;
  let pageCount = 0;
  const MAX_PAGES = 100; // Safety limit to prevent infinite loops
  
  try {
    while (hasNextPage && pageCount < MAX_PAGES) {
      pageCount++;
      let endpoint;
      
      // IMPORTANT: When using page_info, we cannot include other query params
      // The page_info cursor already contains all the filter information
      if (pageInfo) {
        endpoint = `orders.json?limit=250&page_info=${pageInfo}`;
      } else {
        // First request uses the full query params
        endpoint = `orders.json?${queryParams}`;
      }
      
      const url = `https://${SHOPIFY_SHOP_NAME}/admin/api/2024-01/${endpoint}`;
      const response = await axios.get(url, {
        headers: {
          'X-Shopify-Access-Token': SHOPIFY_ADMIN_ACCESS_TOKEN,
          'Content-Type': 'application/json'
        },
        timeout: 60000 // 60 second timeout for large requests
      });
      
      const orders = response.data.orders || [];
      allOrders = allOrders.concat(orders);
      
      console.log(`Page ${pageCount}: Fetched ${orders.length} orders, total so far: ${allOrders.length}`);
      
      // Log detailed order information for debugging
      if (orders.length > 0 && pageCount === 1) {
        console.log('First page order details:');
        orders.slice(0, 5).forEach(order => {
          console.log(`  ${order.name}: fulfillment=${order.fulfillment_status}, financial=${order.financial_status}, status=${order.status}, created=${order.created_at}`);
        });
        console.log(`Query used: ${endpoint}`);
        console.log(`Total orders in response: ${orders.length}`);
        console.log(`All order numbers on this page: ${orders.map(o => o.name).join(', ')}`);
      }
      
      // Check for pagination link in response headers
      const linkHeader = response.headers.link;
      if (linkHeader && linkHeader.includes('rel="next"')) {
        // Extract page_info from the next link
        const nextLinkMatch = linkHeader.match(/<[^>]*page_info=([^>&]+)[^>]*>; rel="next"/);
        if (nextLinkMatch && nextLinkMatch[1]) {
          pageInfo = nextLinkMatch[1];
        } else {
          hasNextPage = false;
        }
      } else {
        hasNextPage = false;
      }
      
      // Safety check: if we got less than 250 orders, we're done
      if (orders.length < 250) {
        hasNextPage = false;
      }
      
      // Add small delay between requests to avoid rate limiting
      if (hasNextPage && pageCount < MAX_PAGES) {
        await new Promise(resolve => setTimeout(resolve, 500)); // 500ms delay
      }
    }
    
    if (pageCount >= MAX_PAGES) {
      console.warn(`⚠️ Reached maximum page limit (${MAX_PAGES}). Total orders: ${allOrders.length}`);
    }
    
    console.log(`✅ Total orders fetched: ${allOrders.length} across ${pageCount} page(s)`);
    return allOrders;
  } catch (error) {
    console.error('Error fetching orders with pagination:', error.response?.data || error.message);
    
    // Provide more specific error messages
    if (error.code === 'ECONNABORTED') {
      throw new Error('Request timeout - try selecting a smaller date range');
    } else if (error.response?.status === 429) {
      throw new Error('Shopify API rate limit exceeded - please wait a moment and try again');
    } else if (error.response?.status === 401) {
      throw new Error('Authentication failed - check your Shopify credentials');
    } else {
      throw new Error(error.response?.data?.errors || error.message || 'Failed to fetch orders from Shopify');
    }
  }
}

// Transform Shopify order data to WooCommerce CSV format
function transformOrderToWooFormat(order, lineItem) {
  const orderNumber = (order.name || order.id).toString().replace('#', '');
  // Use Egypt timezone (UTC+2) for correct local time display
  const orderDate = moment(order.created_at).utcOffset('+02:00').format('YYYY-MM-DD HH:mm');
  const paidDate = order.processed_at ? moment(order.processed_at).utcOffset('+02:00').format('YYYY-MM-DD HH:mm') : '';
  
  return {
    "Order Number": orderNumber,
    "Order Status": order.financial_status || 'pending',
    "Order Date": orderDate,
    "Customer Note": order.note || '',
    "First Name (Billing)": order.billing_address?.first_name || '',
    "Last Name (Billing)": order.billing_address?.last_name || '',
    "Company (Billing)": order.billing_address?.company || '',
    "Address 1&2 (Billing)": `${order.billing_address?.address1 || ''} ${order.billing_address?.address2 || ''}`.trim(),
    "City (Billing)": order.billing_address?.city || '',
    "State Code (Billing)": order.billing_address?.province_code || '',
    "Postcode (Billing)": (order.billing_address?.zip || '').replace(/\s/g, ''),
    "Country Code (Billing)": order.billing_address?.country_code || '',
    "Email (Shipping)": order.shipping_address?.email || order.email || '',
    "Paid Date": paidDate,
    "Phone (Shipping)": (order.shipping_address?.phone || order.phone || '').replace(/\s/g, ''),
    "Phone (Billing)": (order.billing_address?.phone || order.phone || '').replace(/\s/g, ''),
    "First Name (Shipping)": order.shipping_address?.first_name || order.billing_address?.first_name || '',
    "Last Name (Shipping)": order.shipping_address?.last_name || order.billing_address?.last_name || '',
    "Address 1&2 (Shipping)": `${order.shipping_address?.address1 || order.billing_address?.address1 || ''} ${order.shipping_address?.address2 || order.billing_address?.address2 || ''}`.trim(),
    "City (Shipping)": order.shipping_address?.city || order.billing_address?.city || '',
    "State Code (Shipping)": order.shipping_address?.province_code || order.billing_address?.province_code || '',
    "Postcode (Shipping)": (order.shipping_address?.zip || order.billing_address?.zip || '').replace(/\s/g, ''),
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
    "Coupon Code": order.discount_codes?.map(dc => dc.code.replace(/\s/g, '')).join(',') || '',
    "Discount Amount": lineItem.total_discount || '0',
    "Discount Amount Tax": '0',
    "City Code": (order.billing_address?.city || '').replace(/\s/g, ''),
    "Carrier Code": ''
  };
}

// API endpoint for dashboard table view
app.get('/api/orders', async (req, res) => {
  try {
    const { created_at_min, created_at_max, status } = req.query;
    
    console.log('Dashboard API - Request params:', { created_at_min, created_at_max, status });
    
    // Fetch ALL orders - explicitly request all statuses
    // status=any includes open, closed, cancelled, and archived orders
    let queryParams = 'limit=250&status=any';
    if (created_at_min) queryParams += `&created_at_min=${created_at_min}`;
    if (created_at_max) queryParams += `&created_at_max=${created_at_max}`;
    
    console.log('Dashboard API - Query params:', queryParams);
    
    // Use pagination to fetch ALL orders
    const orders = await fetchAllOrders(queryParams);
    
    console.log(`Dashboard API - Returning ${orders.length} orders`);
    
    res.json({
      success: true,
      count: orders.length,
      orders: orders
    });
  } catch (error) {
    console.error('API error:', error);
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

// Export orders endpoint
app.post('/export-orders', async (req, res) => {
  try {
    const { startDate, endDate, status } = req.body;
    
    // Fetch ALL orders - explicitly request all statuses
    // status=any includes open, closed, cancelled, and archived orders
    let queryParams = 'limit=250&status=any';
    if (startDate) {
      const startISO = new Date(startDate).toISOString();
      queryParams += `&created_at_min=${startISO}`;
    }
    if (endDate) {
      // Set to end of day
      const endDateObj = new Date(endDate);
      endDateObj.setHours(23, 59, 59, 999);
      const endISO = endDateObj.toISOString();
      queryParams += `&created_at_max=${endISO}`;
    }
    if (status && status !== 'any') queryParams += `&status=${status}`;
    
    console.log('CSV Export - Query params:', queryParams);
    
    // Use pagination to fetch ALL orders
    const orders = await fetchAllOrders(queryParams);
    
    console.log(`CSV Export - Found ${orders.length} orders`);
    
    if (!orders || orders.length === 0) {
      return res.status(404).json({ error: 'No orders found' });
    }
    
    const csvData = [];
    for (const order of orders) {
      // Skip orders without line items
      if (!order.line_items || order.line_items.length === 0) {
        console.log(`CSV Export - Skipping order ${order.name} - no line items`);
        continue;
      }
      
      for (const lineItem of order.line_items) {
        csvData.push(transformOrderToWooFormat(order, lineItem));
      }
    }
    
    console.log(`CSV Export - Generated ${csvData.length} rows from ${orders.length} orders`);
    
    const csvWriter = createObjectCsvWriter({
      path: 'orders_export.csv',
      encoding: 'utf8',
      append: false,
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
    let csvContent = fs.readFileSync('orders_export.csv', 'utf8');
    
    // Add UTF-8 BOM for Excel to properly display Arabic characters
    const BOM = '\uFEFF';
    csvContent = BOM + csvContent;
    
    res.setHeader('Content-Type', 'text/csv; charset=utf-8');
    res.setHeader('Content-Disposition', 'attachment; filename="orders_export.csv"');
    res.send(csvContent);
    
  } catch (error) {
    console.error('Export error:', error);
    res.status(500).json({ error: error.message });
  }
});

// Export orders to Excel endpoint
app.post('/export-orders-excel', async (req, res) => {
  try {
    const { startDate, endDate, status } = req.body;
    
    // Fetch ALL orders - explicitly request all statuses
    // status=any includes open, closed, cancelled, and archived orders
    let queryParams = 'limit=250&status=any';
    if (startDate) {
      const startISO = new Date(startDate).toISOString();
      queryParams += `&created_at_min=${startISO}`;
    }
    if (endDate) {
      // Set to end of day
      const endDateObj = new Date(endDate);
      endDateObj.setHours(23, 59, 59, 999);
      const endISO = endDateObj.toISOString();
      queryParams += `&created_at_max=${endISO}`;
    }
    if (status && status !== 'any') queryParams += `&status=${status}`;
    
    console.log('Excel Export - Query params:', queryParams);
    
    // Use pagination to fetch ALL orders
    const orders = await fetchAllOrders(queryParams);
    
    console.log(`Excel Export - Found ${orders.length} orders`);
    
    if (!orders || orders.length === 0) {
      console.log('Excel Export - No orders found, returning 404');
      return res.status(404).json({ error: 'No orders found' });
    }
    
    // Create workbook and worksheet
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Orders');
    
    // Define columns (same order as WOO EXAMPLE.csv)
    worksheet.columns = [
      { header: 'Order Number', key: 'orderNumber', width: 15 },
      { header: 'Order Status', key: 'orderStatus', width: 12 },
      { header: 'Order Date', key: 'orderDate', width: 18 },
      { header: 'Customer Note', key: 'customerNote', width: 30 },
      { header: 'First Name (Billing)', key: 'firstNameBilling', width: 18 },
      { header: 'Last Name (Billing)', key: 'lastNameBilling', width: 18 },
      { header: 'Company (Billing)', key: 'companyBilling', width: 20 },
      { header: 'Address 1&2 (Billing)', key: 'addressBilling', width: 35 },
      { header: 'City (Billing)', key: 'cityBilling', width: 18 },
      { header: 'State Code (Billing)', key: 'stateBilling', width: 15 },
      { header: 'Postcode (Billing)', key: 'postcodeBilling', width: 12 },
      { header: 'Country Code (Billing)', key: 'countryBilling', width: 15 },
      { header: 'Email (Shipping)', key: 'emailShipping', width: 25 },
      { header: 'Paid Date', key: 'paidDate', width: 18 },
      { header: 'Phone (Shipping)', key: 'phoneShipping', width: 18 },
      { header: 'Phone (Billing)', key: 'phoneBilling', width: 18 },
      { header: 'First Name (Shipping)', key: 'firstNameShipping', width: 18 },
      { header: 'Last Name (Shipping)', key: 'lastNameShipping', width: 18 },
      { header: 'Address 1&2 (Shipping)', key: 'addressShipping', width: 35 },
      { header: 'City (Shipping)', key: 'cityShipping', width: 18 },
      { header: 'State Code (Shipping)', key: 'stateShipping', width: 15 },
      { header: 'Postcode (Shipping)', key: 'postcodeShipping', width: 12 },
      { header: 'Country Code (Shipping)', key: 'countryShipping', width: 15 },
      { header: 'Payment Method Title', key: 'paymentMethod', width: 20 },
      { header: 'Cart Discount Amount', key: 'cartDiscount', width: 18 },
      { header: 'Order Subtotal Amount', key: 'subtotal', width: 18 },
      { header: 'Shipping Method Title', key: 'shippingMethod', width: 20 },
      { header: 'Order Shipping Amount', key: 'shippingAmount', width: 18 },
      { header: 'Order Refund Amount', key: 'refundAmount', width: 18 },
      { header: 'Order Total Amount', key: 'totalAmount', width: 18 },
      { header: 'Order Total Tax Amount', key: 'taxAmount', width: 18 },
      { header: 'SKU', key: 'sku', width: 20 },
      { header: 'Item #', key: 'itemNumber', width: 15 },
      { header: 'Item Name', key: 'itemName', width: 40 },
      { header: 'Quantity (- Refund)', key: 'quantity', width: 15 },
      { header: 'Item Cost', key: 'itemCost', width: 12 },
      { header: 'Coupon Code', key: 'couponCode', width: 15 },
      { header: 'Discount Amount', key: 'discountAmount', width: 15 },
      { header: 'Discount Amount Tax', key: 'discountTax', width: 15 },
      { header: 'City Code', key: 'cityCode', width: 12 },
      { header: 'Carrier Code', key: 'carrierCode', width: 12 }
    ];
    
    // Style header row
    worksheet.getRow(1).font = { bold: true };
    worksheet.getRow(1).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFE0E0E0' }
    };
    
    // Add data rows
    let rowCount = 0;
    for (const order of orders) {
      // Skip orders without line items
      if (!order.line_items || order.line_items.length === 0) {
        console.log(`Excel Export - Skipping order ${order.name} - no line items`);
        continue;
      }
      
      for (const lineItem of order.line_items) {
        const rowData = transformOrderToWooFormat(order, lineItem);
        worksheet.addRow({
          orderNumber: rowData['Order Number'],
          orderStatus: rowData['Order Status'],
          orderDate: rowData['Order Date'],
          customerNote: rowData['Customer Note'],
          firstNameBilling: rowData['First Name (Billing)'],
          lastNameBilling: rowData['Last Name (Billing)'],
          companyBilling: rowData['Company (Billing)'],
          addressBilling: rowData['Address 1&2 (Billing)'],
          cityBilling: rowData['City (Billing)'],
          stateBilling: rowData['State Code (Billing)'],
          postcodeBilling: rowData['Postcode (Billing)'],
          countryBilling: rowData['Country Code (Billing)'],
          emailShipping: rowData['Email (Shipping)'],
          paidDate: rowData['Paid Date'],
          phoneShipping: rowData['Phone (Shipping)'],
          phoneBilling: rowData['Phone (Billing)'],
          firstNameShipping: rowData['First Name (Shipping)'],
          lastNameShipping: rowData['Last Name (Shipping)'],
          addressShipping: rowData['Address 1&2 (Shipping)'],
          cityShipping: rowData['City (Shipping)'],
          stateShipping: rowData['State Code (Shipping)'],
          postcodeShipping: rowData['Postcode (Shipping)'],
          countryShipping: rowData['Country Code (Shipping)'],
          paymentMethod: rowData['Payment Method Title'],
          cartDiscount: rowData['Cart Discount Amount'],
          subtotal: rowData['Order Subtotal Amount'],
          shippingMethod: rowData['Shipping Method Title'],
          shippingAmount: rowData['Order Shipping Amount'],
          refundAmount: rowData['Order Refund Amount'],
          totalAmount: rowData['Order Total Amount'],
          taxAmount: rowData['Order Total Tax Amount'],
          sku: rowData['SKU'],
          itemNumber: rowData['Item #'],
          itemName: rowData['Item Name'],
          quantity: rowData['Quantity (- Refund)'],
          itemCost: rowData['Item Cost'],
          couponCode: rowData['Coupon Code'],
          discountAmount: rowData['Discount Amount'],
          discountTax: rowData['Discount Amount Tax'],
          cityCode: rowData['City Code'],
          carrierCode: rowData['Carrier Code']
        });
        rowCount++;
      }
    }
    
    console.log(`Excel Export - Generated ${rowCount} rows from ${orders.length} orders`);
    
    // Generate Excel file
    const buffer = await workbook.xlsx.writeBuffer();
    
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="orders_export.xlsx"');
    res.send(buffer);
    
  } catch (error) {
    console.error('Excel export error:', error);
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
