/**
 * Secure Storage Utilities
 * Provides encryption/decryption for sensitive data like API keys
 * Uses Web Crypto API with AES-GCM encryption
 */

import { logger } from './logger';

// Type references for Web Crypto API and Office.js
/// <reference types="@types/office-js" />
/// <reference types="@types/office-runtime" />

// Crypto interface for TypeScript
declare const crypto: {
  subtle: SubtleCrypto;
  getRandomValues: <T extends Uint8Array | Uint32Array>(array: T) => T;
};

// Generate a cryptographic key from a password using PBKDF2
async function deriveKey(password: string, salt: Uint8Array): Promise<CryptoKey> {
  const encoder = new TextEncoder();
  const keyMaterial = await crypto.subtle.importKey(
    'raw',
    encoder.encode(password),
    'PBKDF2',
    false,
    ['deriveKey']
  );

  return crypto.subtle.deriveKey(
    {
      name: 'PBKDF2',
      salt: salt as unknown as BufferSource,
      iterations: 100000,
      hash: 'SHA-256'
    },
    keyMaterial,
    { name: 'AES-GCM', length: 256 },
    false,
    ['encrypt', 'decrypt']
  );
}

// Generate a random salt
function generateSalt(): Uint8Array {
  return crypto.getRandomValues(new Uint8Array(16));
}

// Convert ArrayBuffer to Base64
function arrayBufferToBase64(buffer: ArrayBuffer | ArrayBufferLike): string {
  const bytes = new Uint8Array(buffer as ArrayBuffer);
  let binary = '';
  for (let i = 0; i < bytes.byteLength; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  return btoa(binary);
}

// Convert Base64 to ArrayBuffer
function base64ToArrayBuffer(base64: string): ArrayBuffer {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes.buffer;
}

/**
 * Encrypt a string value using AES-GCM
 * @param value - The plaintext string to encrypt
 * @param password - Optional password (uses device fingerprint if not provided)
 * @returns Encrypted string in format: salt:iv:ciphertext (all base64)
 */
export async function encrypt(value: string, password?: string): Promise<string> {
  const encoder = new TextEncoder();
  const data = encoder.encode(value);
  
  // Use device-specific password or provided password
  const encryptionKey = password || await getDeviceFingerprint();
  
  // Generate salt and derive key
  const salt = generateSalt();
  const key = await deriveKey(encryptionKey, salt);
  
  // Generate IV
  const iv = crypto.getRandomValues(new Uint8Array(12));
  
  // Encrypt
  const ciphertext = await crypto.subtle.encrypt(
    { name: 'AES-GCM', iv },
    key,
    data
  );
  
  // Return as salt:iv:ciphertext (all base64)
  return [
    arrayBufferToBase64(salt.buffer),
    arrayBufferToBase64(iv.buffer),
    arrayBufferToBase64(ciphertext)
  ].join(':');
}

/**
 * Decrypt an encrypted string
 * @param encrypted - The encrypted string in format: salt:iv:ciphertext
 * @param password - Optional password (must match the one used for encryption)
 * @returns Decrypted plaintext string
 */
export async function decrypt(encrypted: string, password?: string): Promise<string> {
  const [saltB64, ivB64, ciphertextB64] = encrypted.split(':');
  
  if (!saltB64 || !ivB64 || !ciphertextB64) {
    throw new Error('Invalid encrypted data format');
  }
  
  // Use device-specific password or provided password
  const encryptionKey = password || await getDeviceFingerprint();
  
  // Decode components
  const salt = new Uint8Array(base64ToArrayBuffer(saltB64));
  const iv = new Uint8Array(base64ToArrayBuffer(ivB64));
  const ciphertext = base64ToArrayBuffer(ciphertextB64);
  
  // Derive key
  const key = await deriveKey(encryptionKey, salt);
  
  // Decrypt
  const decrypted = await crypto.subtle.decrypt(
    { name: 'AES-GCM', iv },
    key,
    ciphertext
  );
  
  const decoder = new TextDecoder();
  return decoder.decode(decrypted);
}

/**
 * Generate a device-specific fingerprint for encryption
 * Uses a combination of browser and Office context identifiers
 */
async function getDeviceFingerprint(): Promise<string> {
  // Try to get Office context identifier
  let officeId = 'unknown';
  try {
    if (typeof Office !== 'undefined' && Office.context) {
      officeId = Office.context.document.url || 'office-document';
    }
  } catch {
    // Office context not available
  }
  
  // Combine with browser identifiers
  const components = [
    navigator.userAgent,
    navigator.language,
    screen.width + 'x' + screen.height,
    new Date().getTimezoneOffset().toString(),
    officeId
  ];
  
  // Hash the components
  const encoder = new TextEncoder();
  const data = encoder.encode(components.join('|'));
  const hashBuffer = await crypto.subtle.digest('SHA-256', data);
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  
  return hashArray.map(b => b.toString(16).padStart(2, '0')).join('').substring(0, 32);
}

/**
 * Secure storage interface for sensitive data
 */
export const secureStorage = {
  /**
   * Store an API key securely
   * @param keyName - Storage key name
   * @param value - The value to store
   */
  async store(keyName: string, value: string): Promise<void> {
    try {
      const encrypted = await encrypt(value);
      
      // Store in Office.settings if available (more secure)
      if (typeof Office !== 'undefined' && Office.context?.document?.settings) {
        Office.context.document.settings.set(keyName, encrypted);
        Office.context.document.settings.saveAsync();
      }
      
      // Also store in localStorage as fallback
      localStorage.setItem(`secure_${keyName}`, encrypted);
    } catch (error) {
      logger.error('Failed to store secure value', { error, keyName });
      throw new Error('Failed to securely store value');
    }
  },

  /**
   * Retrieve a securely stored value
   * @param keyName - Storage key name
   * @returns The decrypted value or null if not found
   */
  async retrieve(keyName: string): Promise<string | null> {
    try {
      let encrypted: string | null = null;
      
      // Try Office.settings first
      if (typeof Office !== 'undefined' && Office.context?.document?.settings) {
        encrypted = Office.context.document.settings.get(keyName) as string | null;
      }
      
      // Fall back to localStorage
      if (!encrypted) {
        encrypted = localStorage.getItem(`secure_${keyName}`);
      }
      
      if (!encrypted) {
        return null;
      }
      
      return await decrypt(encrypted);
    } catch (error) {
      logger.error('Failed to retrieve secure value', { error, keyName });
      return null;
    }
  },

  /**
   * Remove a securely stored value
   * @param keyName - Storage key name
   */
  async remove(keyName: string): Promise<void> {
    try {
      // Remove from Office.settings
      if (typeof Office !== 'undefined' && Office.context?.document?.settings) {
        Office.context.document.settings.remove(keyName);
        Office.context.document.settings.saveAsync();
      }
      
      // Remove from localStorage
      localStorage.removeItem(`secure_${keyName}`);
    } catch (error) {
      logger.error('Failed to remove secure value', { error, keyName });
    }
  },

  /**
   * Check if a value exists
   * @param keyName - Storage key name
   */
  async exists(keyName: string): Promise<boolean> {
    const value = await this.retrieve(keyName);
    return value !== null;
  }
};

export default {
  encrypt,
  decrypt,
  secureStorage
};