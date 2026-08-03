# Google Sheets live API

1. Open <https://script.google.com/home/projects/create> while signed in as the dashboard owner.
2. Replace the standalone project's `Code.gs` with the repository's `Code.gs`.
3. Select **Deploy > New deployment > Web app**.
4. Set **Execute as** to `Me` and **Who has access** to `Anyone`.
5. Authorize and deploy, then copy the `/exec` URL into `docs/data-config.js`.

The API reads the sheet directly by spreadsheet ID and uses a five-minute cache. Use a standalone project so its ownership does not inherit from the source spreadsheet.
