# ════════════════════════════════════════════════════════════════════════
#  OPTION 3 — GITHUB ACTIONS (Fully automated, free, runs in the cloud)
#
#  SETUP STEPS:
#  1. Create a FREE GitHub account at github.com
#  2. Create a new repository called "salesforce-job-alerts"
#  3. Create this file at: .github/workflows/job_alert.yml
#  4. Upload job_alert_runner.py to the root of the repo
#  5. Add Secrets in GitHub repo Settings > Secrets > Actions:
#       APIFY_API_TOKEN   → your Apify token
#       GMAIL_ADDRESS     → sukalyani20@gmail.com
#       GMAIL_APP_PASSWORD → your Gmail App Password
#       RECIPIENT_EMAIL   → sukalyani20@gmail.com
#  6. Push to GitHub — it will run automatically every 6 hours! ✅
# ════════════════════════════════════════════════════════════════════════

name: 🔍 Salesforce QA Job Alert

on:
  schedule:
    - cron: '0 */6 * * *'   # Every 6 hours (0:00, 6:00, 12:00, 18:00 UTC)
  workflow_dispatch:          # Also allows manual trigger from GitHub UI

jobs:
  run-job-alert:
    runs-on: ubuntu-latest
    timeout-minutes: 30

    steps:
      - name: 📥 Checkout repository
        uses: actions/checkout@v4

      - name: 🐍 Set up Python
        uses: actions/setup-python@v5
        with:
          python-version: '3.11'

      - name: 📦 Install dependencies
        run: |
          pip install apify-client openpyxl requests

      - name: 🔍 Run Job Alert Script
        env:
          APIFY_API_TOKEN:    ${{ secrets.APIFY_API_TOKEN }}
          GMAIL_ADDRESS:      ${{ secrets.GMAIL_ADDRESS }}
          GMAIL_APP_PASSWORD: ${{ secrets.GMAIL_APP_PASSWORD }}
          RECIPIENT_EMAIL:    ${{ secrets.RECIPIENT_EMAIL }}
        run: python job_alert_runner.py

      - name: 📊 Upload Excel as Artifact (backup)
        if: always()
        uses: actions/upload-artifact@v4
        with:
          name: job-alert-${{ github.run_number }}
          path: "*.xlsx"
          retention-days: 7
