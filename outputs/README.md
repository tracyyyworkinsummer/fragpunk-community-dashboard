# FragPunk Creator Community Dashboard

Public static dashboard for FragPunk creator/KOC operations.

## Deploy With GitHub Pages

1. Create a public GitHub repository.
2. Upload the contents of this `outputs` folder to the repository root.
3. In GitHub, open **Settings > Pages**.
4. Set **Build and deployment** to **Deploy from a branch**.
5. Choose branch `main` and folder `/root`.
6. Open the GitHub Pages URL after the deployment finishes.

The dashboard reads `dashboard-data.json` as a fast fallback and is structured to poll the public Google Sheet from the browser.
