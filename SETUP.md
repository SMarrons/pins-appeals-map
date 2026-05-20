# PINS Appeals Map — Azure Setup Guide

Complete setup takes about 30–45 minutes.
All services used are free tier or near-zero cost (~£0.01/month).

---

## What you're deploying

```
pins-appeals/
├── function-app/              ← Azure Function (Python) — runs every Monday 6am
│   ├── pins_downloader/
│   │   ├── __init__.py        ← main script: download, parse, geocode, upload
│   │   └── function.json      ← timer trigger config
│   ├── host.json
│   ├── requirements.txt
│   └── local.settings.json    ← local dev only (not deployed)
└── static-web-app/
    └── index.html             ← the map (update DATA_URL before deploying)
```

---

## Step 1 — Prerequisites

Install these tools on your machine:

```bash
# Azure CLI
# Windows: https://aka.ms/installazurecliwindows
# macOS:
brew install azure-cli

# Azure Functions Core Tools (v4)
npm install -g azure-functions-core-tools@4 --unsafe-perm true

# Python 3.11
# https://www.python.org/downloads/

# Git (if not already installed)
# https://git-scm.com/
```

Log in to Azure:
```bash
az login
```

---

## Step 2 — Create Azure resources

Run these commands in order. Replace `yourname` with something unique (lowercase, no spaces).

```bash
# Variables — change these
RESOURCE_GROUP="pins-appeals-rg"
LOCATION="uksouth"
STORAGE_ACCOUNT="pinsappealsYOURNAME"     # must be globally unique, 3-24 lowercase chars
FUNCTION_APP="pins-downloader-YOURNAME"    # must be globally unique
CONTAINER_NAME="pins-data"

# 1. Create resource group
az group create \
  --name $RESOURCE_GROUP \
  --location $LOCATION

# 2. Create storage account (used for both the Function runtime AND the JSON data)
az storage account create \
  --name $STORAGE_ACCOUNT \
  --resource-group $RESOURCE_GROUP \
  --location $LOCATION \
  --sku Standard_LRS \
  --allow-blob-public-access true

# 3. Get the connection string (save this — you'll need it twice)
az storage account show-connection-string \
  --name $STORAGE_ACCOUNT \
  --resource-group $RESOURCE_GROUP \
  --query connectionString \
  --output tsv

# 4. Create the blob container with public read access
az storage container create \
  --name $CONTAINER_NAME \
  --account-name $STORAGE_ACCOUNT \
  --public-access blob

# 5. Enable CORS on the storage account so the browser can fetch the JSON
az storage cors add \
  --account-name $STORAGE_ACCOUNT \
  --services b \
  --methods GET HEAD OPTIONS \
  --origins "*" \
  --allowed-headers "*" \
  --exposed-headers "*" \
  --max-age 3600

# 6. Create the Function App (Python 3.11, consumption plan = free tier)
az functionapp create \
  --resource-group $RESOURCE_GROUP \
  --consumption-plan-location $LOCATION \
  --runtime python \
  --runtime-version 3.11 \
  --functions-version 4 \
  --name $FUNCTION_APP \
  --storage-account $STORAGE_ACCOUNT \
  --os-type Linux

# 7. Set the environment variables the Function needs
#    Replace YOUR_CONNECTION_STRING with the output from step 3
az functionapp config appsettings set \
  --name $FUNCTION_APP \
  --resource-group $RESOURCE_GROUP \
  --settings \
    "STORAGE_CONNECTION_STRING=YOUR_CONNECTION_STRING" \
    "STORAGE_CONTAINER=pins-data"
```

---

## Step 3 — Deploy the Function

```bash
cd pins-appeals/function-app

# Deploy to Azure
func azure functionapp publish $FUNCTION_APP --python
```

You should see output ending with:
```
Deployment successful.
```

---

## Step 4 — Run the Function now (first data load)

Don't wait until Monday — trigger it manually:

```bash
# Option A: via Azure CLI
az rest \
  --method post \
  --uri "https://management.azure.com/subscriptions/$(az account show --query id -o tsv)/resourceGroups/$RESOURCE_GROUP/providers/Microsoft.Web/sites/$FUNCTION_APP/functions/pins_downloader/invoke?api-version=2022-03-01"

# Option B: via Azure Portal
# Go to Function App → Functions → pins_downloader → Code + Test → Test/Run
```

Check the logs to confirm it worked:
```bash
func azure functionapp logstream $FUNCTION_APP
```

You should see lines like:
```
INFO: Downloading PINS data from https://assets.publishing...
INFO: Downloaded 22,800,000 bytes
INFO: Parsed 45,231 records (127 skipped/ungeocodeable)
INFO: Uploaded 8,432,100 bytes → https://pinsappeals....blob.core.windows.net/pins-data/decisions.json
INFO: SUCCESS — 45,231 decisions written
```

---

## Step 5 — Get your blob URL and update the map

```bash
# Print your JSON blob URL
echo "https://${STORAGE_ACCOUNT}.blob.core.windows.net/${CONTAINER_NAME}/decisions.json"
```

Open `static-web-app/index.html` and replace line 3 of the script:

```js
// BEFORE:
const DATA_URL = "https://YOUR_STORAGE_ACCOUNT.blob.core.windows.net/pins-data/decisions.json";

// AFTER (example):
const DATA_URL = "https://pinsappealsjohnsmith.blob.core.windows.net/pins-data/decisions.json";
```

---

## Step 6 — Deploy the Static Web App

```bash
# Create the Static Web App (free tier)
az staticwebapp create \
  --name "pins-appeals-map" \
  --resource-group $RESOURCE_GROUP \
  --location "westeurope" \
  --sku Free

# Get the deployment token
DEPLOY_TOKEN=$(az staticwebapp secrets list \
  --name "pins-appeals-map" \
  --resource-group $RESOURCE_GROUP \
  --query "properties.apiKey" \
  --output tsv)

# Deploy the static site using the SWA CLI
npm install -g @azure/static-web-apps-cli
swa deploy ./static-web-app \
  --deployment-token $DEPLOY_TOKEN \
  --env production
```

Your app URL will be printed, e.g.:
```
https://lemon-sky-0a1b2c3d.azurestaticapps.net
```

---

## Step 7 — Verify everything works end-to-end

1. Open your Static Web App URL in a browser
2. The loading spinner should appear briefly, then the map fills with real PINS decision pins
3. The "Data age" indicator in the top bar shows when data was last refreshed
4. Click any pin to see the full decision detail popup

---

## Automatic weekly updates

The Function is already scheduled. It runs every Monday at 06:00 UTC via this cron expression:

```
0 0 6 * * 1
```

To change the schedule, edit `function-app/pins_downloader/function.json`:
```json
"schedule": "0 0 6 * * 1"   ← Monday 6am UTC
"schedule": "0 0 7 * * 1"   ← Monday 7am UTC  
"schedule": "0 0 6 * * 3"   ← Wednesday 6am UTC
```
Then redeploy: `func azure functionapp publish $FUNCTION_APP --python`

---

## Monitoring & alerts

Set up an email alert if the Function fails:

```bash
# Create action group (replace with your email)
az monitor action-group create \
  --resource-group $RESOURCE_GROUP \
  --name "pins-alerts" \
  --short-name "pinsalert" \
  --action email admin YOUR_EMAIL@example.com

# Create alert rule for Function failures
az monitor metrics alert create \
  --name "pins-function-failures" \
  --resource-group $RESOURCE_GROUP \
  --scopes "/subscriptions/$(az account show --query id -o tsv)/resourceGroups/$RESOURCE_GROUP/providers/Microsoft.Web/sites/$FUNCTION_APP" \
  --condition "count FunctionExecutionUnits > 0 where Result includes 'Failure'" \
  --action "pins-alerts" \
  --description "Alert when PINS downloader fails"
```

---

## Estimated costs

| Service | Usage | Monthly cost |
|---|---|---|
| Azure Functions | 52 executions/year, ~5 min each | £0.00 (free tier: 1M executions) |
| Azure Blob Storage | ~10MB JSON, read by N users | ~£0.01 |
| Azure Static Web Apps | HTML/JS hosting | £0.00 (free tier) |
| **Total** | | **~£0.01/month** |

---

## Troubleshooting

**"CORS error" in browser console**
→ Re-run the `az storage cors add` command from Step 2

**"HTTP 403" fetching the JSON**  
→ Check the container is set to public access:
```bash
az storage container set-permission --name pins-data --account-name $STORAGE_ACCOUNT --public-access blob
```

**Function times out (>10 min)**  
→ The PINS file is large (~22MB). Increase timeout in `host.json`:
```json
"functionTimeout": "00:15:00"
```

**"openpyxl not found" on first deploy**  
→ Ensure `requirements.txt` is in the `function-app/` root (not inside `pins_downloader/`)

**LPA not geocoding (records skipped)**  
→ Check the logs for which LPA names aren't matching. Add aliases to the `LPA_COORDS` dict in `__init__.py` or the fuzzy matcher will handle most variants automatically.
