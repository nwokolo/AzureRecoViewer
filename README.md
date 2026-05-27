# RecoViewer

Web app for Azure benefit recommendation analysis:

- Select billing account and billing profile (or subscription/resource group scope).
- Filter by benefit scope, lookback period, term, and SKU.
- Retrieve recommendations from Cost Management Benefit Recommendations API.
- Chart hourly usage cost and overlay selected recommendation commitment.
- Select recommendation option and overlay mode (commitment line vs effective cost).

## Prerequisites

- Node.js 20+
- Microsoft Entra app registration for SPA login
- Azure RBAC permissions for Cost Management and Billing readers on target scope

## Setup

1. Install backend dependencies:
   - `cd backend`
   - `npm install`
2. Install frontend dependencies:
   - `cd ../frontend`
   - `npm install`
3. Configure environment files:
   - Copy `backend/.env.example` to `backend/.env`
   - Copy `frontend/.env.example` to `frontend/.env`
4. Configure frontend Entra settings in `frontend/.env`:
   - `VITE_AZURE_CLIENT_ID` = SPA app client ID
   - `VITE_AZURE_TENANT_ID` = tenant ID or `common`
   - `VITE_AZURE_REDIRECT_URI` = local frontend URL

## Run

1. Start backend:
   - `cd backend`
   - `npm run dev`
2. Start frontend:
   - `cd frontend`
   - `npm run dev`
3. Open the frontend URL and paste an ARM bearer token.
3. Sign in with Microsoft Entra from the UI.

## Notes

- The app now uses Entra sign-in and silent ARM token acquisition.
- Backend retries throttled requests (429/503) with bounded backoff.
- Benefit recommendations pagination is supported via `nextLink` traversal.

## Validation

- Backend typecheck: `cd backend && npm run typecheck`
- Backend tests: `cd backend && npm test`
- Frontend build: `cd frontend && npm run build`
