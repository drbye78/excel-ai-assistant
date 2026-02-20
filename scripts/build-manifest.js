/**
 * Build script to replace localhost URLs in manifest.xml with production URLs
 * Run this before production builds
 */

const fs = require('fs');
const path = require('path');
require('dotenv').config({ path: path.resolve(__dirname, '../.env') });

const manifestPath = path.resolve(__dirname, '../manifest.xml');
const manifestDevPath = path.resolve(__dirname, '../manifest.dev.xml');

// Get production URL from environment or use default
const prodUrl = process.env.PRODUCTION_URL || process.env.MANIFEST_URL || 'https://localhost:3000';
const devUrl = 'https://localhost:3000';

// Ensure URL ends with no trailing slash
const cleanProdUrl = prodUrl.replace(/\/$/, '');
const cleanDevUrl = devUrl.replace(/\/$/, '');

function buildManifest(isProduction = false) {
  const sourcePath = isProduction ? manifestPath : (fs.existsSync(manifestDevPath) ? manifestDevPath : manifestPath);
  const targetPath = manifestPath;
  
  if (!fs.existsSync(sourcePath)) {
    console.error(`Manifest file not found: ${sourcePath}`);
    process.exit(1);
  }

  let manifestContent = fs.readFileSync(sourcePath, 'utf8');
  
  if (isProduction) {
    // Replace all localhost URLs with production URL
    manifestContent = manifestContent.replace(
      new RegExp(cleanDevUrl.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g'),
      cleanProdUrl
    );
    
    // Update support URL if it's a placeholder
    if (manifestContent.includes('yourusername')) {
      const repoUrl = process.env.REPOSITORY_URL || 'https://github.com/yourusername/excel-ai-assistant';
      manifestContent = manifestContent.replace(
        /https:\/\/github\.com\/yourusername\/excel-ai-assistant/g,
        repoUrl
      );
    }
    
    // Update provider name if it's generic
    if (manifestContent.includes('<ProviderName>AI Assistant Team</ProviderName>')) {
      const providerName = process.env.PROVIDER_NAME || 'AI Assistant Team';
      manifestContent = manifestContent.replace(
        /<ProviderName>AI Assistant Team<\/ProviderName>/,
        `<ProviderName>${providerName}</ProviderName>`
      );
    }
    
    console.log(`✓ Built manifest.xml with production URL: ${cleanProdUrl}`);
  } else {
    console.log(`✓ Using development manifest.xml`);
  }
  
  fs.writeFileSync(targetPath, manifestContent, 'utf8');
}

// Run if called directly
if (require.main === module) {
  const isProduction = process.argv.includes('--production') || process.env.NODE_ENV === 'production';
  buildManifest(isProduction);
}

module.exports = { buildManifest };
