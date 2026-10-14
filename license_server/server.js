/**
 * Bill Scrapper Licensing & Device-Limiting Server
 * Node.js / Express API backed by Supabase PostgreSQL
 */

require('dotenv').config();
const crypto = require('crypto');
const express = require('express');
const cors = require('cors');
const { createClient } = require('@supabase/supabase-js');

const app = express();
const PORT = process.env.PORT || 3000;

// Enable CORS and JSON parsing
app.use(cors());
app.use(express.json());

// Initialize Supabase Client
const SUPABASE_URL = process.env.SUPABASE_URL;
const SUPABASE_SERVICE_ROLE_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY;
const ADMIN_SECRET = process.env.ADMIN_SECRET;

if (!SUPABASE_URL || !SUPABASE_SERVICE_ROLE_KEY) {
  console.error('ERROR: SUPABASE_URL and SUPABASE_SERVICE_ROLE_KEY must be set in environment variables.');
  process.exit(1);
}

// Service role key allows bypassing RLS safely from the secure backend
const supabase = createClient(SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY, {
  auth: { persistSession: false, autoRefreshToken: false }
});

// Plan mapping configurations
const PLAN_CONFIG = {
  '1_month': { duration_days: 30, max_devices: 1, label: '1 Month' },
  '3_months': { duration_days: 90, max_devices: 2, label: '3 Months' },
  '6_months': { duration_days: 180, max_devices: 2, label: '6 Months' },
  '1_year': { duration_days: 365, max_devices: 2, label: '1 Year' }
};

/**
 * Generate a random unique Serial Key in format: BS-XXXX-XXXX-XXXX
 * Uses unambiguous uppercase alphanumeric characters (excluding easily confused 0, O, 1, I)
 */
function generateSerialKey() {
  const CHARSET = '23456789ABCDEFGHJKLMNPQRSTUVWXYZ';
  const getChunk = (len = 4) => {
    const bytes = crypto.randomBytes(len);
    let str = '';
    for (let i = 0; i < len; i++) {
      str += CHARSET[bytes[i] % CHARSET.length];
    }
    return str;
  };
  return `BS-${getChunk(4)}-${getChunk(4)}-${getChunk(4)}`;
}

// ====================================================================
// HEALTH CHECKS
// ====================================================================
app.get(['/', '/health'], (req, res) => {
  res.json({
    status: 'ok',
    app: 'Bill Scrapper License Server',
    timestamp: new Date().toISOString()
  });
});

// ====================================================================
// CORE ACTIVATION / VERIFICATION LOGIC
// Handles both POST /api/activate and POST /api/verify
// ====================================================================
async function handleLicenseVerification(req, res, { isExplicitActivation = false } = {}) {
  try {
    const { serial_key, device_id, os_info } = req.body || {};

    if (!serial_key || typeof serial_key !== 'string') {
      return res.status(400).json({
        valid: false,
        status: 'bad_request',
        message: 'Serial key is required.'
      });
    }

    if (!device_id || typeof device_id !== 'string') {
      return res.status(400).json({
        valid: false,
        status: 'bad_request',
        message: 'Hardware device_id is required.'
      });
    }

    const cleanKey = serial_key.trim().toUpperCase();
    const cleanDeviceId = device_id.trim();
    const cleanOsInfo = (os_info && typeof os_info === 'string') ? os_info.trim() : 'Unknown OS';

    // 1. Fetch license from Supabase
    const { data: license, error: fetchErr } = await supabase
      .from('licenses')
      .select('*')
      .eq('serial_key', cleanKey)
      .maybeSingle();

    if (fetchErr) {
      console.error('[DB ERROR] Fetching license failed:', fetchErr);
      return res.status(500).json({
        valid: false,
        status: 'server_error',
        message: 'Database error while checking license.'
      });
    }

    if (!license) {
      return res.status(404).json({
        valid: false,
        status: 'invalid',
        message: 'Invalid serial key. Please check your key and try again.'
      });
    }

    // 2. Check if revoked
    if (license.is_revoked) {
      return res.status(403).json({
        valid: false,
        status: 'revoked',
        message: 'This serial key has been revoked. Please contact support.'
      });
    }

    const now = new Date();

    // 3. First activation: set activated_at and expires_at
    if (!license.activated_at) {
      const activatedAt = now;
      const expiresAt = new Date(activatedAt.getTime() + license.duration_days * 24 * 60 * 60 * 1000);

      const { error: updateErr } = await supabase
        .from('licenses')
        .update({
          activated_at: activatedAt.toISOString(),
          expires_at: expiresAt.toISOString()
        })
        .eq('id', license.id);

      if (updateErr) {
        console.error('[DB ERROR] Activating license timestamp failed:', updateErr);
        return res.status(500).json({
          valid: false,
          status: 'server_error',
          message: 'Failed to record license activation.'
        });
      }

      license.activated_at = activatedAt.toISOString();
      license.expires_at = expiresAt.toISOString();
      console.log(`[LICENSE ACTIVATED] Key: ${cleanKey} | Plan: ${license.plan_type} | Expires: ${license.expires_at}`);
    }

    // 4. Check expiration
    const expiryDate = new Date(license.expires_at);
    if (now > expiryDate) {
      return res.status(403).json({
        valid: false,
        status: 'expired',
        message: 'Your serial key subscription has expired. Please renew your plan.',
        expires_at: license.expires_at
      });
    }

    // 5. Query registered devices for this license
    const { data: devices, error: devErr } = await supabase
      .from('license_devices')
      .select('*')
      .eq('license_id', license.id);

    if (devErr) {
      console.error('[DB ERROR] Fetching devices failed:', devErr);
      return res.status(500).json({
        valid: false,
        status: 'server_error',
        message: 'Failed to verify device registration.'
      });
    }

    const existingDevice = (devices || []).find(d => d.device_id === cleanDeviceId);

    // Case A: Device is already registered to this license
    if (existingDevice) {
      // Refresh last_active timestamp and OS info
      await supabase
        .from('license_devices')
        .update({
          last_active: now.toISOString(),
          os_info: cleanOsInfo
        })
        .eq('id', existingDevice.id);

      return res.json({
        valid: true,
        status: 'active',
        serial_key: license.serial_key,
        plan_type: license.plan_type,
        duration_days: license.duration_days,
        activated_at: license.activated_at,
        expires_at: license.expires_at,
        devices_used: (devices || []).length,
        max_devices: license.max_devices,
        message: 'License verified successfully.'
      });
    }

    // Case B: Device is NOT registered yet -> enforce max_devices limit
    const currentDeviceCount = (devices || []).length;
    if (currentDeviceCount >= license.max_devices) {
      return res.status(403).json({
        valid: false,
        status: 'limit_reached',
        message: `Device limit reached. This key allows up to ${license.max_devices} device(s), and ${currentDeviceCount} device(s) are already registered.`,
        devices_used: currentDeviceCount,
        max_devices: license.max_devices
      });
    }

    // Register the new device
    const { error: insertDevErr } = await supabase
      .from('license_devices')
      .insert({
        license_id: license.id,
        device_id: cleanDeviceId,
        os_info: cleanOsInfo,
        first_seen: now.toISOString(),
        last_active: now.toISOString()
      });

    if (insertDevErr) {
      console.error('[DB ERROR] Registering new device failed:', insertDevErr);
      return res.status(500).json({
        valid: false,
        status: 'server_error',
        message: 'Failed to register this device.'
      });
    }

    console.log(`[DEVICE REGISTERED] Key: ${cleanKey} | Device: ${cleanDeviceId} | Devices: ${currentDeviceCount + 1}/${license.max_devices}`);

    return res.json({
      valid: true,
      status: 'active',
      serial_key: license.serial_key,
      plan_type: license.plan_type,
      duration_days: license.duration_days,
      activated_at: license.activated_at,
      expires_at: license.expires_at,
      devices_used: currentDeviceCount + 1,
      max_devices: license.max_devices,
      message: 'Device registered and license activated successfully.'
    });

  } catch (err) {
    console.error('[UNEXPECTED ERROR] in verify/activate:', err);
    return res.status(500).json({
      valid: false,
      status: 'server_error',
      message: 'An unexpected internal server error occurred.'
    });
  }
}

// POST /api/activate
app.post('/api/activate', (req, res) => handleLicenseVerification(req, res, { isExplicitActivation: true }));

// POST /api/verify
app.post('/api/verify', (req, res) => handleLicenseVerification(req, res, { isExplicitActivation: false }));

// ====================================================================
// ADMIN ENDPOINT: GENERATE NEW SERIAL KEY
// POST /api/admin/generate
// Protected by x-admin-secret header
// ====================================================================
app.post('/api/admin/generate', async (req, res) => {
  try {
    // 1. Verify admin secret
    const incomingSecret = req.headers['x-admin-secret'] || (req.headers.authorization && req.headers.authorization.replace(/^Bearer\s+/i, ''));

    if (!ADMIN_SECRET || !incomingSecret || incomingSecret !== ADMIN_SECRET) {
      return res.status(401).json({
        success: false,
        error: 'Unauthorized: Invalid or missing x-admin-secret header.'
      });
    }

    // 2. Validate plan_type
    const { plan_type } = req.body || {};
    if (!plan_type || !PLAN_CONFIG[plan_type]) {
      return res.status(400).json({
        success: false,
        error: `Invalid plan_type. Supported values: ${Object.keys(PLAN_CONFIG).join(', ')}`
      });
    }

    const { duration_days, max_devices, label } = PLAN_CONFIG[plan_type];

    // 3. Generate unique serial key (retry in the ultra-rare event of collision)
    let serial_key = null;
    let attempts = 0;

    while (!serial_key && attempts < 5) {
      attempts++;
      const candidateKey = generateSerialKey();

      const { data, error } = await supabase
        .from('licenses')
        .insert({
          serial_key: candidateKey,
          plan_type,
          duration_days,
          max_devices,
          is_revoked: false
        })
        .select('*')
        .single();

      if (!error && data) {
        serial_key = data.serial_key;
        console.log(`[KEY GENERATED] ${candidateKey} | Plan: ${plan_type} (${duration_days} days, max ${max_devices} dev)`);
        return res.status(201).json({
          success: true,
          serial_key: data.serial_key,
          plan_type: data.plan_type,
          plan_label: label,
          duration_days: data.duration_days,
          max_devices: data.max_devices,
          created_at: data.created_at
        });
      }

      // If duplicate key error, try again
      if (error && error.code !== '23505') {
        console.error('[DB ERROR] Generating license failed:', error);
        return res.status(500).json({
          success: false,
          error: 'Database error creating serial key.'
        });
      }
    }

    return res.status(500).json({
      success: false,
      error: 'Failed to generate a unique serial key after multiple attempts.'
    });

  } catch (err) {
    console.error('[UNEXPECTED ERROR] in /api/admin/generate:', err);
    return res.status(500).json({
      success: false,
      error: 'Internal server error while generating key.'
    });
  }
});

// ====================================================================
// SECURE AUTO-UPDATES (Keeps GitHub Repository 100% Private!)
// ====================================================================

const GITHUB_REPO = process.env.GITHUB_REPO || 'naqasht128/BillScrapper';
const GITHUB_TOKEN = process.env.GITHUB_TOKEN || '';

/**
 * GET /api/updates/latest
 * Returns latest version metadata for desktop app without exposing GitHub repo
 */
app.get('/api/updates/latest', async (req, res) => {
  try {
    // 1. If GITHUB_TOKEN is configured in Vercel, query GitHub dynamically
    if (GITHUB_TOKEN) {
      try {
        const ghUrl = `https://api.github.com/repos/${GITHUB_REPO}/releases/latest`;
        const ghRes = await fetch(ghUrl, {
          headers: {
            'User-Agent': 'BillScrapper-LicenseServer',
            'Authorization': `Bearer ${GITHUB_TOKEN}`,
            'Accept': 'application/vnd.github.v3+json'
          }
        });

        if (ghRes.ok) {
          const data = await ghRes.json();
          const tag = (data.tag_name || '').replace(/^v/, '');
          const assets = data.assets || [];

          // Prefer Setup Installer first, then fallback to portable .exe
          const setupAsset = assets.find(a => a.name.toLowerCase().includes('setup') && a.name.toLowerCase().endsWith('.exe'))
            || assets.find(a => a.name.toLowerCase().endsWith('.exe'));

          if (setupAsset) {
            const host = req.get('host');
            const proto = req.headers['x-forwarded-proto'] || 'https';
            return res.json({
              version: tag,
              file_name: setupAsset.name,
              download_url: `${proto}://${host}/api/updates/download?asset_id=${setupAsset.id}`,
              release_notes: data.body || 'New update available with performance enhancements.'
            });
          }
        }
      } catch (ghErr) {
        console.warn('[GH FETCH WARN] Failed to fetch from GitHub API:', ghErr.message);
      }
    }

    // 2. Fallback: Check Supabase table 'app_updates' if present
    try {
      const { data: updateRecord } = await supabase
        .from('app_updates')
        .select('*')
        .order('created_at', { ascending: false })
        .limit(1)
        .maybeSingle();

      if (updateRecord && updateRecord.version) {
        return res.json({
          version: updateRecord.version.replace(/^v/, ''),
          file_name: updateRecord.file_name || 'BillScrapper-Setup.exe',
          download_url: updateRecord.download_url,
          release_notes: updateRecord.release_notes || 'New update available.'
        });
      }
    } catch (dbErr) {
      // Table may not exist yet, continue to default fallback
    }

    // 3. Fallback to environment variables
    const fallbackVersion = (process.env.LATEST_VERSION || '1.2.0').replace(/^v/, '');
    const fallbackUrl = process.env.LATEST_DOWNLOAD_URL || '';
    return res.json({
      version: fallbackVersion,
      file_name: 'BillScrapper-Setup.exe',
      download_url: fallbackUrl,
      release_notes: 'Latest version of Bill Scrapper.'
    });

  } catch (err) {
    console.error('[UPDATES ERROR]', err);
    return res.status(500).json({ error: 'Failed to check for updates' });
  }
});

/**
 * GET /api/updates/download
 * Proxies/streams the installer file from private GitHub Release securely
 */
app.get('/api/updates/download', async (req, res) => {
  try {
    const { asset_id } = req.query;

    if (asset_id && GITHUB_TOKEN) {
      const assetUrl = `https://api.github.com/repos/${GITHUB_REPO}/releases/assets/${asset_id}`;
      const ghRes = await fetch(assetUrl, {
        headers: {
          'User-Agent': 'BillScrapper-LicenseServer',
          'Authorization': `Bearer ${GITHUB_TOKEN}`,
          'Accept': 'application/octet-stream'
        },
        redirect: 'follow'
      });

      if (!ghRes.ok) {
        return res.status(ghRes.status).send('Failed to fetch update asset from storage.');
      }

      const contentType = ghRes.headers.get('content-type') || 'application/octet-stream';
      const contentLength = ghRes.headers.get('content-length');
      res.setHeader('Content-Type', contentType);
      if (contentLength) res.setHeader('Content-Length', contentLength);
      res.setHeader('Content-Disposition', 'attachment; filename="BillScrapper-Setup.exe"');

      const arrayBuffer = await ghRes.arrayBuffer();
      return res.send(Buffer.from(arrayBuffer));
    }

    // Direct URL redirect fallback
    const directUrl = process.env.LATEST_DOWNLOAD_URL;
    if (directUrl) {
      return res.redirect(directUrl);
    }

    return res.status(404).send('Download file not found.');
  } catch (err) {
    console.error('[DOWNLOAD PROXY ERROR]', err);
    return res.status(500).send('Error downloading update file.');
  }
});

/**
 * POST /api/admin/set-update
 * Allows admin to register or update the latest version link in Supabase
 */
app.post('/api/admin/set-update', async (req, res) => {
  try {
    const adminSecret = req.headers['x-admin-secret'];
    if (!ADMIN_SECRET || adminSecret !== ADMIN_SECRET) {
      return res.status(401).json({ error: 'Unauthorized: Invalid admin secret' });
    }

    const { version, download_url, file_name, release_notes } = req.body || {};
    if (!version) {
      return res.status(400).json({ error: 'Version is required.' });
    }

    const cleanVer = String(version).trim().replace(/^v/, '');

    const { error } = await supabase
      .from('app_updates')
      .upsert({
        version: cleanVer,
        download_url: download_url || '',
        file_name: file_name || 'BillScrapper-Setup.exe',
        release_notes: release_notes || `Release v${cleanVer}`,
        updated_at: new Date().toISOString()
      }, { onConflict: 'version' });

    if (error) {
      console.error('[DB ERROR] set-update failed:', error);
      return res.status(500).json({ error: error.message });
    }

    return res.json({ success: true, message: `Update v${cleanVer} registered successfully!` });
  } catch (err) {
    console.error('[ADMIN SET UPDATE ERROR]', err);
    return res.status(500).json({ error: 'Internal server error' });
  }
});

// Start listening locally / standalone
if (require.main === module) {
  app.listen(PORT, () => {
    console.log(`====================================================`);
    console.log(`🚀 Bill Scrapper License Server running on port ${PORT}`);
    console.log(`   Health check: http://localhost:${PORT}/health`);
    console.log(`====================================================`);
  });
}

// Export for serverless environments (Vercel)
module.exports = app;

