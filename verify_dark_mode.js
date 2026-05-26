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

  // We need to mock google.script.run
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
                      timestamp: new Date(Date.now() - 3000000).toString(), // ~50 mins ago
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
                      timestamp: new Date(Date.now() - 100000000).toString(), // > 24h ago
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
                      timestamp: new Date(Date.now() - 200000000).toString(), // > 48h ago
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
                  handler([]);
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

  // Toggle dark mode
  await page.evaluate(() => {
    document.body.classList.add('dark-mode');
  });

  await page.setViewportSize({ width: 1280, height: 800 });
  await page.waitForTimeout(2000); // Wait for rendering
  await page.screenshot({ path: '/home/jules/verification/screenshots/dark_mode_verification.png', fullPage: true });
  await page.waitForTimeout(1000);

  await context.close();
  await browser.close();
})();
