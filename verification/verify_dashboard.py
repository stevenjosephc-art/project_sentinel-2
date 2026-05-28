
from playwright.sync_api import sync_playwright
import os

def run_verification(page):
    abs_path = os.path.abspath("Dashboard.html")
    page.goto(f"file://{abs_path}")

    # Improved mock for google.script.run
    page.evaluate("""
        const mockData = {
            session: { email: 'sme_user@google.com', isSME: true },
            openCases: [
                {
                    rowIdx: 2, viewers: ['viewer1', 'viewer2'], timestamp: new Date().toString(),
                    submitter: 'agent1@google.com', ldap: 'agent1_ldap', caseId: '1-1234567890123',
                    symptom: 'Payment Issue', detailedIssue: 'Customer cannot pay', reason: 'High value',
                    channel: 'Chat', team: 'Team Steven', caseLink: 'http://case/1', status: 'Open',
                    claimedBy: '', isMine: false
                },
                {
                    rowIdx: 3, viewers: [], timestamp: new Date().toString(),
                    submitter: 'agent2@google.com', ldap: 'agent2_ldap', caseId: '1-9876543210987',
                    symptom: 'Refund Request', detailedIssue: 'Accidental purchase', reason: 'Over limit',
                    channel: 'Email', team: 'Team Gerry', caseLink: 'http://case/2', status: 'In Progress',
                    claimedBy: 'sme_user@google.com', isMine: true
                }
            ],
            resolvedCases: [
                {
                    rowIdx: 10, timestamp: new Date().toString(), submitter: 'agent3@google.com',
                    ldap: 'agent3_ldap', caseId: '1-0000000000000', symptom: 'Technical Issue',
                    detailedIssue: 'App crash', reason: 'Escalated', channel: 'Phone', team: 'Team Khent',
                    caseLink: 'http://case/3', claimedAt: new Date().toString(),
                    resolutionTime: new Date().toString(), handledBy: 'sme_user@google.com',
                    remarks: 'Fixed the issue via PPP education.',
                    resolutionType: 'Process > Speed of resolution > Consult'
                }
            ],
            schedule: [
                {
                    email: 'sme_user@google.com', ldap: 'sme_user', role: 'SME',
                    days: [
                        { label: 'Today', dateLabel: 'Jan-01', value: '09:00:00 AM', status: 'on' },
                        { label: 'Tomorrow', dateLabel: 'Jan-02', value: '09:00:00 AM', status: 'on' },
                        { label: 'Wed Jan 03', dateLabel: 'Jan-03', value: 'OFF', status: 'off' }
                    ]
                }
            ],
            webAppUrl: 'https://mock-url'
        };

        function createMockProxy(handler = null) {
            const proxy = {
                withSuccessHandler: (h) => createMockProxy(h),
                withFailureHandler: (h) => createMockProxy(handler),
                getSessionAndRole: () => handler && handler(mockData.session),
                getOpenCases: () => handler && handler(mockData.openCases),
                getResolvedCases: () => handler && handler(mockData.resolvedCases),
                getSchedule: () => handler && handler(mockData.schedule),
                getWebAppUrl: () => handler && handler(mockData.webAppUrl),
                pingCase: () => {}
            };
            return proxy;
        }

        window.google = { script: { run: createMockProxy() } };

        // Delay the onload to ensure mocks are ready
        if (window.onload) window.onload();
    """)

    page.wait_for_timeout(2000)
    page.screenshot(path="verification/screenshots/dashboard_active.png")

    page.click('[data-tab="resolved"]')
    page.wait_for_timeout(1000)
    page.screenshot(path="verification/screenshots/dashboard_resolved.png")

    page.click('[data-tab="metrics"]')
    page.wait_for_timeout(2000)
    page.screenshot(path="verification/screenshots/dashboard_metrics.png")

    page.click('[data-tab="schedule"]')
    page.wait_for_timeout(1000)
    page.screenshot(path="verification/screenshots/dashboard_schedule.png")

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
