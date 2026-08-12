const fs = require('fs');
const jsdom = require('jsdom');
const { JSDOM } = jsdom;

const htmlContent = fs.readFileSync('C:\\Users\\GoldenTech\\inventory-nour\\index.html', 'utf8');

// Mock data
const mockData = {
  inventory: [
    { name: 'Item A', balance: 50, category: 'خامات' },
    { name: 'Item B', balance: 10, category: 'خامات' },
    { name: 'Item C', balance: 0.5, category: 'خامات' }, // Should be excluded from balance KPIs
    { name: 'Item D', balance: 0, category: 'خامات' }  // Should be excluded from balance KPIs
  ],
  stats: {
    total_items: 4,
    total_balance: 60.5,
    low_stock: 1,
    total_transactions: 10
  },
  invoices: {
    items: [
      {
        name: 'Item A',
        total_out: 100,
        total_revenue: 1000,
        daily_customers: {
          '2026-08-10': { 'Cust 1': 50, 'Cust 2': 50 }
        }
      },
      {
        name: 'Item B',
        total_out: 200,
        total_revenue: 2000,
        daily_customers: {
          '2026-08-10': { 'Cust 1': 150, 'Cust 3': 50 }
        }
      }
    ]
  }
};

const dom = new JSDOM(htmlContent, { runScripts: 'dangerously' });
const window = dom.window;

// Override globals in the JSDOM context so they execute without errors
window.EMBEDDED_DATA = mockData;

// Wait for scripts to execute and populate functions
setTimeout(() => {
  try {
    const { renderKPIs, filterByDate, renderCustomersTab, renderBalanceTab } = window;
    let passed = 0;
    let failed = 0;

    function assert(condition, message) {
      if (condition) {
        console.log(`✅ PASS: ${message}`);
        passed++;
      } else {
        console.error(`❌ FAIL: ${message}`);
        failed++;
      }
    }

    console.log('\n━━━ 1. Balance KPIs UI Logic Test ━━━');
    const balanceKPIs = renderKPIs(mockData, 'balance');
    // Balance tab valid items are those >= 1.
    // Item A (50), Item B (10). Total = 2 items.
    // Item C (0.5), Item D (0) are excluded.
    assert(balanceKPIs.includes('>' + 2 + '<'), 'Total valid items count is 2');
    assert(balanceKPIs.includes('>' + 60 + '<'), 'Total balance is 60 (50 + 10)');
    assert(balanceKPIs.includes('>' + 1 + '<'), 'Low stock count is 1 (Item B has 10)');

    console.log('\n━━━ 2. filterByDate (all_customers) Test ━━━');
    const filtered = filterByDate(mockData.invoices.items);
    assert(filtered.length === 2, 'Filtered items count is 2');
    assert(filtered[0].all_customers !== undefined, 'all_customers populated for item A');
    assert(filtered[0].all_customers.find(c => c.name === 'Cust 1').qty === 50, 'Cust 1 qty in item A is correct');

    console.log('\n━━━ 3. Customers KPI Test ━━━');
    const custKPIs = renderKPIs(mockData, 'customers');
    // Cust 1 (200), Cust 2 (50), Cust 3 (50). Total = 3 active customers.
    assert(custKPIs.includes('>' + 3 + '<'), 'Total active customers is 3');
    assert(custKPIs.includes('>Cust 1<'), 'Top customer is Cust 1');

    console.log(`\n════════════════════════════════════════`);
    console.log(`  Total: ${passed + failed}`);
    console.log(`  Passed: ${passed}`);
    console.log(`  Failed: ${failed}`);
    if (failed === 0) {
      console.log(`  🎉 All tests passed!`);
    } else {
      console.log(`  🚨 Some tests failed.`);
    }

  } catch (err) {
    console.error('Error during testing:', err);
  }
}, 500);
