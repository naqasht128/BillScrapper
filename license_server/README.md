# Bill Scrapper License Server

A Node.js & Express licensing, subscription, and device-limiting API backed by Supabase PostgreSQL, ready for 1-click deployment on Render.

---

## 1. Supabase Setup (PostgreSQL Database)

1. Create a free project at [supabase.com](https://supabase.com).
2. Go to your **SQL Editor** tab in the Supabase Dashboard.
3. Open `license_server/schema.sql` from this repository, copy its contents, paste them into the SQL Editor, and click **Run**.
4. Go to **Project Settings** -> **API**:
   - Copy **Project URL** (`SUPABASE_URL`)
   - Copy **service_role secret key** (`SUPABASE_SERVICE_ROLE_KEY`) *(⚠️ Never expose this key in public client apps)*

---

## 2. Deploying to Render

1. Push your repository to GitHub.
2. Log into [render.com](https://render.com) and click **New +** -> **Web Service**.
3. Connect your repository and configure:
   - **Root Directory**: `license_server`
   - **Runtime**: `Node`
   - **Build Command**: `npm install`
   - **Start Command**: `node server.js`
   - **Plan**: `Free`
4. In the **Environment Variables** section, add:
   - `SUPABASE_URL` = your Supabase Project URL
   - `SUPABASE_SERVICE_ROLE_KEY` = your Supabase service_role secret key
   - `ADMIN_SECRET` = any strong random secret string (e.g. `MySuperSecretKey2026`)
   - `PORT` = `10000` (Render default)
5. Click **Deploy Web Service**.
6. Copy your public URL (e.g. `https://billscrapper-api.onrender.com`).

---

## 3. API Endpoints

### Health Check
- `GET /health`
- Response: `{"status": "ok", "app": "Bill Scrapper License Server", "timestamp": "..."}`

### Verify or Activate License
- `POST /api/verify` or `POST /api/activate`
- Header: `Content-Type: application/json`
- Body:
```json
{
  "serial_key": "BS-XXXX-XXXX-XXXX",
  "device_id": "unique-hardware-sha256-hash",
  "os_info": "Windows 11 (AMD64)"
}
```
- Success (200):
```json
{
  "valid": true,
  "status": "active",
  "serial_key": "BS-XXXX-XXXX-XXXX",
  "plan_type": "1_month",
  "duration_days": 30,
  "expires_at": "2026-11-14T10:00:00.000Z",
  "devices_used": 1,
  "max_devices": 1
}
```
- Rejections (403/404):
  - `status: "invalid"`: Serial key not found.
  - `status: "expired"`: Subscription has elapsed.
  - `status: "limit_reached"`: Already activated on maximum allowed devices.
  - `status: "revoked"`: Key explicitly disabled by administrator.

### Admin: Generate New Key
- `POST /api/admin/generate`
- Header: `x-admin-secret: <your ADMIN_SECRET>`
- Body:
```json
{
  "plan_type": "1_month"
}
```
*Supported `plan_type` values:*
- `1_month`: 30 Days, 1 Device (Rs. 300)
- `3_months`: 90 Days, 2 Devices (Rs. 700)
- `6_months`: 180 Days, 2 Devices (Rs. 1200)
- `1_year`: 365 Days, 2 Devices (Rs. 2000)

---

## 4. Local Testing
```bash
cd license_server
npm install
npm run dev
```
Test health check at `http://localhost:3000/health`.
