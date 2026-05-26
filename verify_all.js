const { chromium } = require('playwright');
const fs = require('fs');

(async () => {
  const browser = await chromium.launch();
  const context = await browser.newContext({
    recordVideo: {
      dir: '/home/jules/verification/videos',
    }
  });
  const page = await context.newPage();

  const htmlContent = fs.readFileSync('Dashboard.html', 'utf8');

  const mockedHtml = htmlContent.replace('</head>', `
    <script>
      window.google = {
        script: {
          run: {
            withSuccessHandler: function(handler) {
              const runner = {
                getSessionAndRole: function() {
                  handler({
                    email: 'test@google.com',
                    isSME: true
                  });
                },
                getWebAppUrl: function() {
                  handler('https://script.google.com/macros/s/123/exec');
                },
                getOpenCases: function() {
                  handler([
                    {
                      rowIdx: 3,
                      timestamp: new Date(Date.now() - 3000000).toString(),
                      ldap: 'testagent',
                      caseId: '1-2345678901234',
                      symptom: 'Payment issue',
                      reason: 'Blocked account',
                      channel: 'Chat',
                      team: 'Team Steven',
                      status: 'Open'
                    },
                    {
                      rowIdx: 4,
                      timestamp: new Date(Date.now() - 100000000).toString(),
                      ldap: 'agent2',
                      caseId: '2-2345678901234',
                      symptom: 'Refund request',
                      reason: 'User unhappy',
                      channel: 'Email',
                      team: 'Team Gerry',
                      status: 'Open'
                    },
                    {
                      rowIdx: 5,
                      timestamp: new Date(Date.now() - 200000000).toString(),
                      ldap: 'agent3',
                      caseId: '3-2345678901234',
                      symptom: 'Technical error',
                      reason: 'App crash',
                      channel: 'Phone',
                      team: 'Team James',
                      status: 'Open'
                    }
                  ]);
                },
                getResolvedCases: function() {
                  handler([
                    {
                      rowIdx: 10,
                      resolutionTime: new Date(Date.now() - 3600000).toString(),
                      ldap: 'agent1',
                      caseId: 'R-12345',
                      symptom: 'Refund approved',
                      handledBy: 'sme@google.com',
                      channel: 'Email',
                      team: 'Team Steven',
                      resolutionType: 'Process > Speed of resolution > Contact time',
                      remarks: 'Approved after verification.'
                    }
                  ]);
                },
                getSchedule: function() {
                  handler([]);
                }
              };
              return runner;
            }
          }
        }
      };
    </script>
  </head>`);

  await page.setContent(mockedHtml);
  await page.setViewportSize({ width: 1280, height: 800 });

  // Light Mode - Active
  await page.waitForTimeout(1000);
  await page.screenshot({ path: '/home/jules/verification/screenshots/light_active.png', fullPage: true });

  // Dark Mode - Active
  await page.evaluate(() => document.body.classList.add('dark-mode'));
  await page.waitForTimeout(1000);
  await page.screenshot({ path: '/home/jules/verification/screenshots/dark_active.png', fullPage: true });

  // Dark Mode - Resolved
  await page.click('div[data-tab="resolved"]');
  await page.waitForTimeout(1000);
  await page.screenshot({ path: '/home/jules/verification/screenshots/dark_resolved.png', fullPage: true });

  await context.close();
  await browser.close();
})();
