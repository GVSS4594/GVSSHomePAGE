name: Deploy Staging

on:
  push:
    branches:
      - staging

jobs:
  deploy-staging:
    runs-on: ubuntu-latest
    steps:
      - name: Checkout code
        uses: actions/checkout@v4

      - name: Setup Node.js
        uses: actions/setup-node@v4
        with:
          node-version: '22'

      - name: Install Google Apps Script Clasp
        run: npm install -g @google/clasp

      - name: Authenticate Clasp
        run: |
          echo '{"access_token":"${{ secrets.CLASP_ACCESS_TOKEN }}","refresh_token":"${{ secrets.CLASP_REFRESH_TOKEN }}","scope":"https://www.googleapis.com/auth/script.projects https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/script.deployments","token_type":"Bearer","expiry_date":1600000000000}' > ~/.clasprc.json

      - name: Push to Google Apps Script Staging
        run: clasp push --force