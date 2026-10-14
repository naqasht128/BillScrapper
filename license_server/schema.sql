-- ====================================================================
-- Bill Scrapper - Supabase PostgreSQL Database Schema
-- Run this script in the Supabase SQL Editor: Dashboard -> SQL Editor
-- ====================================================================

-- 1. Ensure required extensions are available
CREATE EXTENSION IF NOT EXISTS "pgcrypto";

-- ====================================================================
-- 2. Table: licenses
-- Stores issued serial keys, subscription plans, and validity
-- ====================================================================
CREATE TABLE IF NOT EXISTS public.licenses (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    serial_key TEXT UNIQUE NOT NULL,
    plan_type TEXT NOT NULL CHECK (plan_type IN ('1_month', '3_months', '6_months', '1_year')),
    duration_days INTEGER NOT NULL CHECK (duration_days IN (30, 90, 180, 365)),
    max_devices INTEGER NOT NULL CHECK (max_devices IN (1, 2)),
    activated_at TIMESTAMPTZ,
    expires_at TIMESTAMPTZ,
    is_revoked BOOLEAN NOT NULL DEFAULT FALSE,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- ====================================================================
-- 3. Table: license_devices
-- Tracks hardware devices activated under each serial key
-- ====================================================================
CREATE TABLE IF NOT EXISTS public.license_devices (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    license_id UUID NOT NULL REFERENCES public.licenses(id) ON DELETE CASCADE,
    device_id TEXT NOT NULL,
    os_info TEXT,
    first_seen TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    last_active TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    CONSTRAINT uq_license_device UNIQUE (license_id, device_id)
);

-- ====================================================================
-- 4. Performance Indexes
-- ====================================================================
CREATE INDEX IF NOT EXISTS idx_licenses_serial_key ON public.licenses (serial_key);
CREATE INDEX IF NOT EXISTS idx_licenses_expires_at ON public.licenses (expires_at);
CREATE INDEX IF NOT EXISTS idx_license_devices_license_id ON public.license_devices (license_id);
CREATE INDEX IF NOT EXISTS idx_license_devices_device_id ON public.license_devices (device_id);

-- ====================================================================
-- 5. Row Level Security (RLS) Configuration
-- Enable RLS for database safety. The Node.js Express server connects
-- using the Supabase Service Role Key, which automatically bypasses RLS.
-- Public/anonymous direct access without the backend API is blocked.
-- ====================================================================
ALTER TABLE public.licenses ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.license_devices ENABLE ROW LEVEL SECURITY;

-- Deny anon/public direct access (all requests MUST go through the Render Express API)
DROP POLICY IF EXISTS "Deny anon access on licenses" ON public.licenses;
CREATE POLICY "Deny anon access on licenses" ON public.licenses
    FOR ALL TO anon USING (false);

DROP POLICY IF EXISTS "Deny anon access on license_devices" ON public.license_devices;
CREATE POLICY "Deny anon access on license_devices" ON public.license_devices
    FOR ALL TO anon USING (false);

-- ====================================================================
-- Comments for documentation
-- ====================================================================
COMMENT ON TABLE public.licenses IS 'Serial keys, plans, and activation lifecycles for Bill Scrapper.';
COMMENT ON TABLE public.license_devices IS 'Hardware devices bound to active licenses with device limits.';

-- ====================================================================
-- 3. APP UPDATES TABLE
-- ====================================================================
CREATE TABLE IF NOT EXISTS public.app_updates (
    version VARCHAR(32) PRIMARY KEY,
    download_url TEXT NOT NULL,
    file_name VARCHAR(128) DEFAULT 'BillScrapper-Setup.exe',
    release_notes TEXT,
    created_at TIMESTAMPTZ DEFAULT NOW(),
    updated_at TIMESTAMPTZ DEFAULT NOW()
);

ALTER TABLE public.app_updates ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "Allow anon read on app_updates" ON public.app_updates;
CREATE POLICY "Allow anon read on app_updates" ON public.app_updates
    FOR SELECT TO anon USING (true);
DROP POLICY IF EXISTS "Deny anon modify on app_updates" ON public.app_updates;
CREATE POLICY "Deny anon modify on app_updates" ON public.app_updates
    FOR INSERT TO anon WITH CHECK (false);
