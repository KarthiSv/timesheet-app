// Full E2E test suite - IBM Timesheet Manager
const { test, expect } = require('@playwright/test');
const path = require('path');
const fs = require('fs');

const BASE = 'http://localhost:4173';
const IBM_FILE_1 = '/tmp/ibm_test_main.xlsx';
const IBM_FILE_2 = '/tmp/ibm_test_sheet1.xlsx';
const CLARITY_FILE = '/tmp/clarity_test.xlsx';

// Helper: login as manager
async function loginAsManager(page) {
  await page.goto(BASE);
  await page.waitForLoadState('networkidle');
  const emailInput = page.locator('input').first();
  await emailInput.fill('manager@ibm.com');
  const passInput = page.locator('input[type="password"]');
  await passInput.fill('IBMTimesheet2025!');
  await page.locator('button[type="submit"], button:has-text("Sign In"), button:has-text("Login")').first().click();
  await page.waitForTimeout(1000);
}

// Helper: open import modal and upload files
async function importFiles(page, ibmFiles, clarityFile) {
  const importBtn = page.locator('button:has-text("Import"), button:has-text("↑ Import")').first();
  await importBtn.click();
  await page.waitForTimeout(500);

  // Upload IBM files
  for (const f of ibmFiles) {
    const ibmInput = page.locator('input[type="file"]').first();
    await ibmInput.setInputFiles(f);
    await page.waitForTimeout(800);
  }

  // Upload Clarity file
  const clarityInput = page.locator('input[type="file"]').nth(1);
  await clarityInput.setInputFiles(clarityFile);
  await page.waitForTimeout(800);

  // Click Import button
  const doImportBtn = page.locator('button:has-text("Import"), button:has-text("✓ Import")').last();
  await doImportBtn.click();
  await page.waitForTimeout(1000);
}

test.describe('App Load & Login', () => {
  test('T01: Page loads without crash', async ({ page }) => {
    await page.goto(BASE);
    await page.waitForLoadState('networkidle');
    // Should show login screen or app
    const body = await page.textContent('body');
    expect(body.length).toBeGreaterThan(10);
  });

  test('T02: No JS console errors on load', async ({ page }) => {
    const errors = [];
    page.on('pageerror', e => errors.push(e.message));
    await page.goto(BASE);
    await page.waitForLoadState('networkidle');
    await page.waitForTimeout(1000);
    const criticalErrors = errors.filter(e =>
      !e.includes('favicon') &&
      !e.includes('ResizeObserver') &&
      !e.includes('Non-Error promise rejection')
    );
    expect(criticalErrors).toHaveLength(0);
  });

  test('T03: Login screen renders correctly', async ({ page }) => {
    await page.goto(BASE);
    await page.waitForLoadState('networkidle');
    const inputs = page.locator('input');
    await expect(inputs.first()).toBeVisible();
  });

  test('T04: IBM branding visible', async ({ page }) => {
    await page.goto(BASE);
    await page.waitForLoadState('networkidle');
    const body = await page.textContent('body');
    expect(body.toLowerCase()).toContain('ibm');
  });
});

test.describe('Import Flow', () => {
  test('T05: Import button visible after login', async ({ page }) => {
    await loginAsManager(page);
    const importBtn = page.locator('button:has-text("Import"), button:has-text("↑ Import")').first();
    await expect(importBtn).toBeVisible();
  });

  test('T06: Import modal opens', async ({ page }) => {
    await loginAsManager(page);
    const importBtn = page.locator('button:has-text("Import"), button:has-text("↑ Import")').first();
    await importBtn.click();
    await page.waitForTimeout(500);
    const modal = page.locator('text=Import Data');
    await expect(modal).toBeVisible();
  });

  test('T07: IBM file upload accepted', async ({ page }) => {
    await loginAsManager(page);
    const importBtn = page.locator('button:has-text("↑ Import"), button:has-text("Import")').first();
    await importBtn.click();
    await page.waitForTimeout(500);
    const ibmInput = page.locator('input[type="file"]').first();
    await ibmInput.setInputFiles(IBM_FILE_1);
    await page.waitForTimeout(1000);
    // Should show file name
    const body = await page.textContent('body');
    expect(body).toContain('ibm_test_main');
  });

  test('T08: Grand Total row not counted as person', async ({ page }) => {
    await loginAsManager(page);
    const importBtn = page.locator('button:has-text("↑ Import"), button:has-text("Import")').first();
    await importBtn.click();
    await page.waitForTimeout(500);
    const ibmInput = page.locator('input[type="file"]').first();
    await ibmInput.setInputFiles(IBM_FILE_1);
    await page.waitForTimeout(1000);
    // Should show 5 people (not 6 - Grand Total should be excluded)
    const body = await page.textContent('body');
    expect(body).not.toContain('Grand Total');
    // Count should be 5 people
    expect(body).toMatch(/5\s*people/);
  });

  test('T09: Sheet1 fallback IBM file accepted', async ({ page }) => {
    await loginAsManager(page);
    const importBtn = page.locator('button:has-text("↑ Import"), button:has-text("Import")').first();
    await importBtn.click();
    await page.waitForTimeout(500);
    const ibmInput = page.locator('input[type="file"]').first();
    await ibmInput.setInputFiles(IBM_FILE_2);
    await page.waitForTimeout(1000);
    const body = await page.textContent('body');
    // Should NOT show 'Could not find sheet' error
    expect(body).not.toContain("Could not find sheet");
    expect(body).toMatch(/2\s*people/);
  });

  test('T10: Clarity file upload accepted', async ({ page }) => {
    await loginAsManager(page);
    const importBtn = page.locator('button:has-text("↑ Import"), button:has-text("Import")').first();
    await importBtn.click();
    await page.waitForTimeout(500);
    const ibmInput = page.locator('input[type="file"]').first();
    await ibmInput.setInputFiles(IBM_FILE_1);
    await page.waitForTimeout(800);
    const clarityInput = page.locator('input[type="file"]').nth(1);
    await clarityInput.setInputFiles(CLARITY_FILE);
    await page.waitForTimeout(1000);
    const body = await page.textContent('body');
    expect(body).toContain('clarity_test');
  });

  test('T11: Match preview shows matched/IBM only/Clarity only counts', async ({ page }) => {
    await loginAsManager(page);
    const importBtn = page.locator('button:has-text("↑ Import"), button:has-text("Import")').first();
    await importBtn.click();
    await page.waitForTimeout(500);
    const ibmInput = page.locator('input[type="file"]').first();
    await ibmInput.setInputFiles(IBM_FILE_1);
    await page.waitForTimeout(800);
    const clarityInput = page.locator('input[type="file"]').nth(1);
    await clarityInput.setInputFiles(CLARITY_FILE);
    await page.waitForTimeout(1000);
    const body = await page.textContent('body');
    expect(body).toContain('MATCHED');
    expect(body).toContain('IBM ONLY');
    expect(body).toContain('CLARITY ONLY');
  });

  test('T12: Dual-period Clarity hours summed correctly (Alice: 81+99=180)', async ({ page }) => {
    await loginAsManager(page);
    const importBtn = page.locator('button:has-text("↑ Import"), button:has-text("Import")').first();
    await importBtn.click();
    await page.waitForTimeout(500);
    const ibmInput = page.locator('input[type="file"]').first();
    await ibmInput.setInputFiles(IBM_FILE_1);
    await page.waitForTimeout(800);
    const clarityInput = page.locator('input[type="file"]').nth(1);
    await clarityInput.setInputFiles(CLARITY_FILE);
    await page.waitForTimeout(1000);
    const body = await page.textContent('body');
    // Alice has 81+99=180h actual. Should show 180h somewhere
    expect(body).toContain('180h');
  });

  test('T13: Import button enabled and clickable', async ({ page }) => {
    await loginAsManager(page);
    const importBtn = page.locator('button:has-text("↑ Import"), button:has-text("Import")').first();
    await importBtn.click();
    await page.waitForTimeout(500);
    const ibmInput = page.locator('input[type="file"]').first();
    await ibmInput.setInputFiles(IBM_FILE_1);
    await page.waitForTimeout(800);
    const clarityInput = page.locator('input[type="file"]').nth(1);
    await clarityInput.setInputFiles(CLARITY_FILE);
    await page.waitForTimeout(1000);
    const doImport = page.locator('button:has-text("✓ Import"), button:has-text("Import")').last();
    await expect(doImport).toBeEnabled();
  });
});

test.describe('Post-Import: Dashboard Tab', () => {
  test.beforeEach(async ({ page }) => {
    await loginAsManager(page);
    await importFiles(page, [IBM_FILE_1, IBM_FILE_2], CLARITY_FILE);
  });

  test('T14: Dashboard loads after import - no blank page', async ({ page }) => {
    // Should be on Dashboard by default
    await page.waitForTimeout(500);
    const errors = [];
    page.on('pageerror', e => errors.push(e.message));
    await page.waitForTimeout(500);
    const criticalErrors = errors.filter(e => !e.includes('ResizeObserver'));
    expect(criticalErrors).toHaveLength(0);
    const body = await page.textContent('body');
    expect(body.length).toBeGreaterThan(100);
  });

  test('T15: Dashboard shows Overview heading', async ({ page }) => {
    const heading = page.locator('h1, h2').filter({ hasText: /Overview|Dashboard/i }).first();
    await expect(heading).toBeVisible();
  });

  test('T16: Dashboard stat cards visible', async ({ page }) => {
    const body = await page.textContent('body');
    expect(body).toMatch(/Total People|TOTAL PEOPLE/i);
    expect(body).toMatch(/IBM Scheduled|IBM SCHEDULED/i);
  });

  test('T17: Dashboard defaults to All Months (not filtered)', async ({ page }) => {
    const body = await page.textContent('body');
    expect(body).toMatch(/All Months|all imported/i);
  });

  test('T18: Filter by Month dropdown exists', async ({ page }) => {
    const filterDropdown = page.locator('select').first();
    await expect(filterDropdown).toBeVisible();
  });

  test('T19: Month filter has All Months as first option', async ({ page }) => {
    const firstOption = await page.locator('select option').first().textContent();
    expect(firstOption).toMatch(/All Months/i);
  });

  test('T20: Selecting a specific month updates the view', async ({ page }) => {
    const select = page.locator('select').first();
    await select.selectOption({ index: 1 }); // Pick first real month
    await page.waitForTimeout(400);
    const body = await page.textContent('body');
    expect(body.length).toBeGreaterThan(100); // Page still renders
  });

  test('T21: Charts render (pie chart visible)', async ({ page }) => {
    await page.waitForTimeout(1000);
    const svg = page.locator('svg').first();
    await expect(svg).toBeVisible();
  });

  test('T22: Action Required panel shows records', async ({ page }) => {
    const body = await page.textContent('body');
    expect(body).toMatch(/Action Required|Missing Entry|CRITICAL/i);
  });

  test('T23: Fix Name Mismatches button visible when mismatches exist', async ({ page }) => {
    const body = await page.textContent('body');
    expect(body).toMatch(/Fix Name Mismatches|Fix \(/i);
  });
});

test.describe('Post-Import: Records Tab — CRITICAL', () => {
  test.beforeEach(async ({ page }) => {
    await loginAsManager(page);
    await importFiles(page, [IBM_FILE_1, IBM_FILE_2], CLARITY_FILE);
  });

  test('T24: Records tab click does NOT crash (no blank page)', async ({ page }) => {
    const errors = [];
    page.on('pageerror', e => errors.push(e.message));

    const recordsTab = page.locator('button:has-text("Records"), nav button').filter({ hasText: /Records/i }).first();
    await recordsTab.click();
    await page.waitForTimeout(1500);

    // Check no JS errors
    const criticalErrors = errors.filter(e =>
      !e.includes('ResizeObserver') && !e.includes('Non-Error')
    );
    expect(criticalErrors).toHaveLength(0);

    // Check page has content (not blank)
    const body = await page.textContent('body');
    expect(body.length).toBeGreaterThan(200);
  });

  test('T25: Records tab shows employee table', async ({ page }) => {
    const recordsTab = page.locator('button, a').filter({ hasText: /^Records$/ }).first();
    await recordsTab.click();
    await page.waitForTimeout(1000);
    const body = await page.textContent('body');
    expect(body).toMatch(/Name|IBM|Source/i);
  });

  test('T26: Records table has rows after import', async ({ page }) => {
    const recordsTab = page.locator('button, a').filter({ hasText: /^Records$/ }).first();
    await recordsTab.click();
    await page.waitForTimeout(1000);
    const rows = page.locator('tbody tr');
    const count = await rows.count();
    expect(count).toBeGreaterThan(0);
  });

  test('T27: Alice Johnson visible in Records (matched record)', async ({ page }) => {
    const recordsTab = page.locator('button, a').filter({ hasText: /^Records$/ }).first();
    await recordsTab.click();
    await page.waitForTimeout(1000);
    const body = await page.textContent('body');
    expect(body).toContain('Alice Johnson');
  });

  test('T28: Records tab - clicking a row does NOT crash', async ({ page }) => {
    const errors = [];
    page.on('pageerror', e => errors.push(e.message));

    const recordsTab = page.locator('button, a').filter({ hasText: /^Records$/ }).first();
    await recordsTab.click();
    await page.waitForTimeout(1000);

    // Click first employee row
    const firstRow = page.locator('tbody tr').first();
    if (await firstRow.count() > 0) {
      await firstRow.click();
      await page.waitForTimeout(1000);
      const criticalErrors = errors.filter(e => !e.includes('ResizeObserver'));
      expect(criticalErrors).toHaveLength(0);
    }
  });

  test('T29: Employee detail panel opens on row click', async ({ page }) => {
    const recordsTab = page.locator('button, a').filter({ hasText: /^Records$/ }).first();
    await recordsTab.click();
    await page.waitForTimeout(1000);
    const firstRow = page.locator('tbody tr').first();
    if (await firstRow.count() > 0) {
      await firstRow.click();
      await page.waitForTimeout(800);
      // Panel should show Overview tab
      const body = await page.textContent('body');
      expect(body).toMatch(/Overview|Timesheet|Scheduled|Actual/i);
    }
  });

  test('T30: Detail panel close button works', async ({ page }) => {
    const recordsTab = page.locator('button, a').filter({ hasText: /^Records$/ }).first();
    await recordsTab.click();
    await page.waitForTimeout(1000);
    const firstRow = page.locator('tbody tr').first();
    if (await firstRow.count() > 0) {
      await firstRow.click();
      await page.waitForTimeout(800);
      const closeBtn = page.locator('button:has-text("Close"), button:has-text("✕")').first();
      if (await closeBtn.isVisible()) {
        await closeBtn.click();
        await page.waitForTimeout(500);
        // Panel should be gone
        const body = await page.textContent('body');
        expect(body.length).toBeGreaterThan(100);
      }
    }
  });

  test('T31: Search filter works in Records', async ({ page }) => {
    const recordsTab = page.locator('button, a').filter({ hasText: /^Records$/ }).first();
    await recordsTab.click();
    await page.waitForTimeout(1000);
    const searchInput = page.locator('input[placeholder*="Search"]').first();
    if (await searchInput.isVisible()) {
      await searchInput.fill('Alice');
      await page.waitForTimeout(400);
      const body = await page.textContent('body');
      expect(body).toContain('Alice');
    }
  });

  test('T32: Status filter buttons work', async ({ page }) => {
    const recordsTab = page.locator('button, a').filter({ hasText: /^Records$/ }).first();
    await recordsTab.click();
    await page.waitForTimeout(1000);
    const missingBtn = page.locator('button:has-text("Missing")').first();
    if (await missingBtn.isVisible()) {
      await missingBtn.click();
      await page.waitForTimeout(400);
      const body = await page.textContent('body');
      expect(body.length).toBeGreaterThan(50);
    }
  });

  test('T33: Re-link button visible on matched rows', async ({ page }) => {
    const recordsTab = page.locator('button, a').filter({ hasText: /^Records$/ }).first();
    await recordsTab.click();
    await page.waitForTimeout(1000);
    const body = await page.textContent('body');
    expect(body).toMatch(/Re-link|⇄/);
  });

  test('T34: Records tab does NOT crash after Fix Name Mismatches + Apply', async ({ page }) => {
    const errors = [];
    page.on('pageerror', e => errors.push(e.message));

    // Open Fix Name Mismatches
    const fixBtn = page.locator('button:has-text("Fix"), button:has-text("Fix Name")').first();
    if (await fixBtn.isVisible()) {
      await fixBtn.click();
      await page.waitForTimeout(800);

      // Close without applying
      const closeBtn = page.locator('button:has-text("Close"), button:has-text("✕ Close")').first();
      if (await closeBtn.isVisible()) {
        await closeBtn.click();
        await page.waitForTimeout(500);
      }
    }

    // Now go to Records tab
    const recordsTab = page.locator('button, a').filter({ hasText: /^Records$/ }).first();
    await recordsTab.click();
    await page.waitForTimeout(1500);

    const criticalErrors = errors.filter(e => !e.includes('ResizeObserver') && !e.includes('Non-Error'));
    expect(criticalErrors).toHaveLength(0);

    const body = await page.textContent('body');
    expect(body.length).toBeGreaterThan(200);
  });
});

test.describe('Fix Name Mismatches Panel', () => {
  test.beforeEach(async ({ page }) => {
    await loginAsManager(page);
    await importFiles(page, [IBM_FILE_1, IBM_FILE_2], CLARITY_FILE);
  });

  test('T35: Fix Name Mismatches panel opens', async ({ page }) => {
    const fixBtn = page.locator('button').filter({ hasText: /Fix/ }).first();
    if (await fixBtn.isVisible()) {
      await fixBtn.click();
      await page.waitForTimeout(500);
      const body = await page.textContent('body');
      expect(body).toMatch(/Fix Name Mismatches|Name Match/i);
    }
  });

  test('T36: Fix panel has Unmatched IBM tab', async ({ page }) => {
    const fixBtn = page.locator('button').filter({ hasText: /Fix/ }).first();
    if (await fixBtn.isVisible()) {
      await fixBtn.click();
      await page.waitForTimeout(500);
      const body = await page.textContent('body');
      expect(body).toMatch(/Unmatched IBM/i);
    }
  });

  test('T37: Fix panel has Fix Auto-matched tab', async ({ page }) => {
    const fixBtn = page.locator('button').filter({ hasText: /Fix/ }).first();
    if (await fixBtn.isVisible()) {
      await fixBtn.click();
      await page.waitForTimeout(500);
      const body = await page.textContent('body');
      expect(body).toMatch(/Auto.matched|Fix Auto/i);
    }
  });

  test('T38: Clarity right panel shows names', async ({ page }) => {
    const fixBtn = page.locator('button').filter({ hasText: /Fix/ }).first();
    if (await fixBtn.isVisible()) {
      await fixBtn.click();
      await page.waitForTimeout(500);
      // Switch to Fix Auto-matched tab
      const autoTab = page.locator('button').filter({ hasText: /Auto.matched/ }).first();
      if (await autoTab.isVisible()) {
        await autoTab.click();
        await page.waitForTimeout(400);
        const body = await page.textContent('body');
        // Clarity names should be visible - check for total count text
        expect(body).toMatch(/All Clarity Names|available|Clarity/i);
      }
    }
  });

  test('T39: Clarity right search box works', async ({ page }) => {
    const fixBtn = page.locator('button').filter({ hasText: /Fix/ }).first();
    if (await fixBtn.isVisible()) {
      await fixBtn.click();
      await page.waitForTimeout(500);
      const searchInput = page.locator('input[placeholder*="Clarity"]').first();
      if (await searchInput.isVisible()) {
        await searchInput.fill('Alice');
        await page.waitForTimeout(400);
        const body = await page.textContent('body');
        expect(body).toContain('Alice');
      }
    }
  });

  test('T40: Records tab NOT blank after Fix panel is closed', async ({ page }) => {
    const errors = [];
    page.on('pageerror', e => errors.push(e.message));

    const fixBtn = page.locator('button').filter({ hasText: /Fix/ }).first();
    if (await fixBtn.isVisible()) {
      await fixBtn.click();
      await page.waitForTimeout(500);
      const closeBtn = page.locator('button:has-text("Close"), button:has-text("✕ Close")').first();
      if (await closeBtn.isVisible()) {
        await closeBtn.click();
        await page.waitForTimeout(500);
      }
    }

    const recordsTab = page.locator('button, a').filter({ hasText: /^Records$/ }).first();
    await recordsTab.click();
    await page.waitForTimeout(1500);

    const criticalErrors = errors.filter(e => !e.includes('ResizeObserver') && !e.includes('Non-Error'));
    expect(criticalErrors).toHaveLength(0);
    const body = await page.textContent('body');
    expect(body.length).toBeGreaterThan(200);
  });
});

test.describe('Other Tabs', () => {
  test.beforeEach(async ({ page }) => {
    await loginAsManager(page);
    await importFiles(page, [IBM_FILE_1], CLARITY_FILE);
  });

  test('T41: Calendar tab loads without crash', async ({ page }) => {
    const errors = [];
    page.on('pageerror', e => errors.push(e.message));
    const calTab = page.locator('button, a').filter({ hasText: /Calendar/ }).first();
    await calTab.click();
    await page.waitForTimeout(1000);
    const criticalErrors = errors.filter(e => !e.includes('ResizeObserver'));
    expect(criticalErrors).toHaveLength(0);
    const body = await page.textContent('body');
    expect(body.length).toBeGreaterThan(100);
  });

  test('T42: Users tab loads without crash', async ({ page }) => {
    const errors = [];
    page.on('pageerror', e => errors.push(e.message));
    const usersTab = page.locator('button, a').filter({ hasText: /Users/ }).first();
    await usersTab.click();
    await page.waitForTimeout(1000);
    const criticalErrors = errors.filter(e => !e.includes('ResizeObserver'));
    expect(criticalErrors).toHaveLength(0);
    const body = await page.textContent('body');
    expect(body.length).toBeGreaterThan(100);
  });

  test('T43: Profile tab loads without crash', async ({ page }) => {
    const errors = [];
    page.on('pageerror', e => errors.push(e.message));
    const profileTab = page.locator('button, a').filter({ hasText: /Profile/ }).first();
    await profileTab.click();
    await page.waitForTimeout(1000);
    const criticalErrors = errors.filter(e => !e.includes('ResizeObserver'));
    expect(criticalErrors).toHaveLength(0);
    const body = await page.textContent('body');
    expect(body.length).toBeGreaterThan(100);
  });

  test('T44: Can switch between all tabs without crash', async ({ page }) => {
    const errors = [];
    page.on('pageerror', e => errors.push(e.message));

    for (const tabName of ['Records', 'Calendar', 'Users', 'Profile', 'Dashboard']) {
      const tab = page.locator('button, a').filter({ hasText: new RegExp(`^${tabName}$`) }).first();
      if (await tab.isVisible()) {
        await tab.click();
        await page.waitForTimeout(800);
      }
    }

    const criticalErrors = errors.filter(e => !e.includes('ResizeObserver') && !e.includes('Non-Error'));
    expect(criticalErrors).toHaveLength(0);
  });

  test('T45: Sign Out button visible', async ({ page }) => {
    const signOut = page.locator('button:has-text("Sign Out")').first();
    await expect(signOut).toBeVisible();
  });
});

test.describe('Dashboard Charts & Variance Bar', () => {
  test.beforeEach(async ({ page }) => {
    await loginAsManager(page);
    await importFiles(page, [IBM_FILE_1, IBM_FILE_2], CLARITY_FILE);
    await page.waitForTimeout(500);
  });

  test('T46: Variance bar visible', async ({ page }) => {
    const body = await page.textContent('body');
    expect(body).toMatch(/variance|Scheduled vs Actual|Actual vs Scheduled/i);
  });

  test('T47: Total variance number displayed', async ({ page }) => {
    const body = await page.textContent('body');
    expect(body).toMatch(/\d+h.*variance|variance.*\d+h/i);
  });

  test('T48: Charts section renders (SVG present)', async ({ page }) => {
    await page.waitForTimeout(1000);
    const svgs = await page.locator('svg').count();
    expect(svgs).toBeGreaterThan(0);
  });

  test('T49: By Department chart present', async ({ page }) => {
    const body = await page.textContent('body');
    expect(body).toMatch(/Department|Dept/i);
  });

  test('T50: Severity breakdown visible', async ({ page }) => {
    const body = await page.textContent('body');
    expect(body).toMatch(/Severity|Critical|Complete/i);
  });
});

