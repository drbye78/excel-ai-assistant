/**
 * Input Sanitization Utilities
 * 
 * Prevents formula injection and sanitizes inputs for different contexts
 */

/**
 * Sanitize Excel cell values to prevent formula injection
 * Prefixes formula-like values with single quote to make them text
 * @param value - Value to sanitize
 * @returns Sanitized value safe for Excel cells
 */
export function sanitizeCellValue(value: unknown): string {
  if (value === null || value === undefined) return '';
  
  const str = String(value);
  
  // Prevent formula injection (cells starting with =, +, -, @)
  if (/^[=+\-@]/.test(str)) {
    return `'${str}`;
  }
  
  return str;
}

/**
 * Sanitize VBA macro names to ensure valid VBA syntax
 */
export function sanitizeVbaName(name: string): string {
  return name
    .replace(/^[^a-zA-Z]+/, '')
    .replace(/[^a-zA-Z0-9_]/g, '_')
    .substring(0, 255);
}

/**
 * Sanitize Power Query inputs to prevent SQL-like injection
 */
export function sanitizeQueryInput(input: string): string {
  return input
    .replace(/'/g, "''")
    .replace(/;/g, '')
    .replace(/--/g, '')
    .replace(/\/\*/g, '')
    .replace(/\*\//g, '');
}

/**
 * Sanitize HTML content for display
 */
export function sanitizeHtml(input: string): string {
  const htmlEntities: Record<string, string> = {
    '<': '&lt;',
    '>': '&gt;',
    '&': '&amp;',
    '"': '&quot;',
    "'": '&#x27;',
  };
  
  return input.replace(/[<>&"']/g, char => htmlEntities[char]);
}

/**
 * Sanitize file path for Windows/Unix compatibility
 */
export function sanitizeFilePath(path: string): string {
  return path
    .replace(/[<>:"|?*]/g, '')
    .replace(/\.\./g, '')
    .replace(/\/+/g, '/')
    .trim();
}

/**
 * Validate and sanitize email address
 */
export function sanitizeEmail(email: string): { valid: boolean; sanitized?: string; error?: string } {
  const trimmed = email.trim().toLowerCase();
  
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!emailRegex.test(trimmed)) {
    return { valid: false, error: 'Invalid email format' };
  }
  
  if (/[<>()\\[\];:&]/.test(trimmed)) {
    return { valid: false, error: 'Email contains invalid characters' };
  }
  
  return { valid: true, sanitized: trimmed };
}

/**
 * Sanitize JSON string for safe parsing
 */
export function sanitizeJson(json: string): string {
  return json.replace(/[\x00-\x1F\x7F-\x9F]/g, '');
}

/**
 * Check if a string contains potentially dangerous content
 */
export function isSafeContent(content: string): boolean {
  const dangerousPatterns = [
    /javascript:/i,
    /data:text\/html/i,
    /vbscript:/i,
    /on\w+\s*=/i,
    /<script/i,
    /<iframe/i,
    /<object/i,
    /<embed/i,
  ];
  
  return !dangerousPatterns.some(pattern => pattern.test(content));
}

/**
 * Sanitize SQL-like identifiers (table names, column names)
 */
export function sanitizeIdentifier(identifier: string): string {
  const sanitized = identifier.replace(/[^a-zA-Z0-9_]/g, '');
  
  if (/^[^a-zA-Z]/.test(sanitized)) {
    return `col_${sanitized}`;
  }
  
  return sanitized;
}

/**
 * Sanitize for use in Excel named ranges
 */
export function sanitizeRangeName(name: string): string {
  return name
    .replace(/[^a-zA-Z0-9_]/g, '_')
    .replace(/^[^a-zA-Z]/, 'Range_')
    .substring(0, 255);
}

/**
 * Sanitize cell reference for use in formulas
 */
export function sanitizeCellReference(ref: string): string {
  const sanitized = ref.toUpperCase().replace(/[^A-Z0-9$:]/g, '');
  
  const cellRefPattern = /^\$?[A-Z]+\$?\d+(:\$?[A-Z]+\$?\d+)?$/;
  if (!cellRefPattern.test(sanitized)) {
    throw new Error(`Invalid cell reference: ${ref}`);
  }
  
  return sanitized;
}
