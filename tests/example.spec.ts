import { test, expect } from '@playwright/test';
import 'dotenv/config';


test('test', async ({ page }) => {
  await page.goto('https://depediloilo2a.duckdns.org:12443/depediloilo/sakto/apps/pis/login.php');
  await page.getByRole('link', { name: ' Sign in with Google' }).click();
  await page.getByRole('textbox', { name: 'Email or phone' }).click();
  await page.getByRole('textbox', { name: 'Email or phone' }).fill(process.env.EMAIL);
  await page.getByRole('textbox', { name: 'Email or phone' }).press('Enter');
  await page.getByRole('textbox', { name: 'Enter your password' }).fill(process.env.GOOGLE_PASSWORD);
  await page.getByRole('textbox', { name: 'Enter your password' }).press('Enter');
  await page.getByRole('button', { name: 'Continue' }).click();
  await page.getByRole('textbox', { name: 'Password' }).click();
  await page.getByRole('textbox', { name: 'Password' }).fill(process.env.PIS_PASSWORD);
  await page.getByRole('button', { name: 'Login' }).click();
  await page.getByText('Time Records').click();
  await page.getByRole('link', { name: 'Work Time Attendance' }).click();
  await page.locator('#select2-selectSCHOOLASSIGNMENT-container').click();
  await page.getByRole('textbox').fill('10000');
  await page.getByRole('treeitem', { name: '- Schools Division of Iloilo' }).click();
  await page.getByRole('searchbox', { name: 'Search:' }).click();
  await page.getByRole('searchbox', { name: 'Search:' }).fill('ponte');
  await page.locator('[id="33940"]').getByTitle('Print TimeSheet').click();
  await page.locator('#txtDATERANGEPRINT').click();
  await page.getByRole('cell', { name: '1' }).first().click();
  const context = page.context();

const [newPage] = await Promise.all([
  context.waitForEvent('page'),
  page.locator('#displayPRINTTIMESHEETSDIALOG')
    .getByTitle('Display Monthly Timesheets')
    .click()
]);

await newPage.evaluate(() => {
  const target = document.getElementById('main');

  if (!target) return;

  document.body.innerHTML = '';
  document.body.appendChild(target);
});
await page.waitForTimeout(3000);
await newPage.pdf({ path: 'output.pdf', printBackground: true });
});



