# Azure Deployment Plan

## Status
Planning

## Objective
Deploy RecoViewer to Azure so it is publicly accessible.

## Workspace Mode
MODIFY existing app.

## Selected Deployment Recipe
Azure CLI (no `azd` present).

## Azure Context
- Subscription: `BenefitsPMSub` (`15511700-9115-4a1c-8482-ecfdd4089b53`) detected from current `az` context
- Location: `eastus` (proposed)
- Resource group: `rg-rec-viewer-prod` (proposed)

## Architecture Plan
- Backend: Azure App Service (Linux, Node 20) deployed from `backend`
- Frontend: Azure Storage Static Website deployed from `frontend/dist`
- Frontend API base URL: backend public URL from App Service

## Required Changes
1. `backend/package.json`
- Add `postinstall` script to build TypeScript on App Service deployment.

2. Build/deploy pipeline (CLI-driven)
- Deploy backend with `az webapp up`.
- Build frontend with `VITE_API_BASE_URL` pointing to backend URL.
- Enable static website on Storage Account and upload frontend `dist` files.

## Public Endpoints (Expected)
- Backend: `https://<webapp-name>.azurewebsites.net`
- Frontend: `https://<storage-account>.z13.web.core.windows.net` (region suffix may vary)

## Security/Access Notes
- Backend remains publicly reachable for frontend consumption.
- CORS is currently open in app code; acceptable for initial public deployment.

## Validation Plan
- Hit backend health endpoint.
- Load frontend URL and verify recommendation data retrieval.

## Execution Steps
1. Create resource group.
2. Deploy backend App Service.
3. Retrieve backend public URL.
4. Build frontend with production API URL.
5. Create storage account and enable static website.
6. Upload frontend assets.
7. Verify backend and frontend endpoints.

## Rollback Plan
- Delete created resource group to remove all resources:
  `az group delete --name rg-rec-viewer-prod --yes --no-wait`
