from playwright.sync_api import sync_playwright
import os

def run_verification(page):
    dashboard_path = "file://" + os.path.abspath("Dashboard.html")
    page.goto(dashboard_path)
    page.wait_for_timeout(1000)

    # Use a simpler evaluate block
    page.evaluate("""
        const mockData = {
            getSessionAndRole: {
                email: 'jules@google.com',
                isSME: true
            },
            getOpenCases: [
                {
                    rowIdx: 2,
                    ldap: 'cyrilmark',
                    caseId: '9-3019000040750',
                    symptom: 'Don\\\'t want a purchase',
                    detailedIssue: 'Customer wants a refund because they changed their mind after downloading the app but before using it.',
                    reason: 'Supervisor request',
                    channel: 'Chat',
                    team: 'Team Mary',
                    timestamp: new Date(Date.now() - 32 * 3600000).toString(),
                    status: 'Open',
                    viewers: ['criseldaa']
                }
            ],
            getResolvedCases: [],
            getSchedule: [],
            getWebAppUrl: 'https://script.google.com/test'
        };

        const createProxy = (handler) => {
            return {
                withSuccessHandler: (h) => createProxy(h),
                withFailureHandler: () => createProxy(handler),
                getSessionAndRole: () => { if(handler) handler(mockData.getSessionAndRole); },
                getOpenCases: () => { if(handler) handler(mockData.getOpenCases); },
                getResolvedCases: () => { if(handler) handler(mockData.getResolvedCases); },
                getSchedule: () => { if(handler) handler(mockData.getSchedule); },
                getWebAppUrl: () => { if(handler) handler(mockData.getWebAppUrl); },
                pingCase: () => {}
            };
        };

        window.google = {
            script: {
                run: createProxy()
            }
        };

        window.onload();
    """)
    page.wait_for_timeout(2000)

    # 1. Capture Light Mode Dashboard
    page.screenshot(path="verification/screenshots/dashboard_light.png")
    page.wait_for_timeout(500)

    # 2. Toggle Dark Mode
    page.get_by_title("Toggle Dark Mode").click()
    page.wait_for_timeout(1000)
    page.screenshot(path="verification/screenshots/dashboard_dark.png")

    page.wait_for_timeout(1000)

if __name__ == "__main__":
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(record_video_dir="verification/videos")
        page = context.new_page()
        try:
            run_verification(page)
        finally:
            context.close()
            browser.close()
