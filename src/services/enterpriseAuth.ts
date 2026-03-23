/**
 * Enterprise Authentication Service
 *
 * Provides SSO (Single Sign-On), multi-user support, and enterprise
 * authentication features for the Excel AI Assistant.
 *
 * Features:
 * - SSO integration (Azure AD, Okta, SAML)
 * - Multi-user session management
 * - Role-based access control (RBAC)
 * - User provisioning and deprovisioning
 * - Audit logging
 *
 * @module services/enterpriseAuth
 */

import { notificationManager } from '../utils/notificationManager';
import { logger } from '../utils/logger';
import { PermissionError, AppError } from '../utils/errors';

// ============================================================================
// Type Definitions
// ============================================================================

/** Authentication provider types */
export type AuthProvider = 'azure_ad' | 'okta' | 'saml' | 'oauth2' | 'local';

/** User roles */
export type UserRole = 'admin' | 'power_user' | 'standard_user' | 'guest';

/** User permissions */
export interface UserPermissions {
  canUseAI: boolean;
  canCreateRecipes: boolean;
  canShareContent: boolean;
  canAccessAdminPanel: boolean;
  canManageUsers: boolean;
  canViewAuditLogs: boolean;
  canExportData: boolean;
  maxQueriesPerDay: number;
  maxRecipes: number;
}

/** Enterprise user */
export interface EnterpriseUser {
  id: string;
  email: string;
  displayName: string;
  role: UserRole;
  department?: string;
  permissions: UserPermissions;
  provider: AuthProvider;
  externalId?: string;
  lastLoginAt?: Date;
  createdAt: Date;
  isActive: boolean;
  mfaEnabled: boolean;
  sessionExpiry?: Date;
  passwordHash?: string; // For local authentication
}

/** SSO configuration */
export interface SSOConfig {
  provider: AuthProvider;
  clientId: string;
  clientSecret?: string;
  tenantId?: string;
  authority?: string;
  redirectUri: string;
  scopes: string[];
  metadataUrl?: string;
  certificate?: string;
}

/** Session info */
export interface SessionInfo {
  sessionId: string;
  userId: string;
  startedAt: Date;
  expiresAt: Date;
  ipAddress?: string;
  userAgent?: string;
  isValid: boolean;
}

/** Audit log entry */
export interface AuditLogEntry {
  id: string;
  timestamp: Date;
  userId: string;
  action: string;
  resource: string;
  details: Record<string, any>;
  ipAddress?: string;
  userAgent?: string;
  success: boolean;
  errorMessage?: string;
}

/** Team/Organization */
export interface Organization {
  id: string;
  name: string;
  domain: string;
  ssoConfig?: SSOConfig;
  settings: OrganizationSettings;
  createdAt: Date;
}

/** Organization settings */
export interface OrganizationSettings {
  requireMfa: boolean;
  sessionTimeoutMinutes: number;
  maxUsers: number;
  allowedProviders: AuthProvider[];
  dataRetentionDays: number;
  auditLogRetentionDays: number;
  enableSharing: boolean;
  enableExternalAI: boolean;
}

// ============================================================================
// Role Permissions
// ============================================================================

const ROLE_PERMISSIONS: Record<UserRole, UserPermissions> = {
  admin: {
    canUseAI: true,
    canCreateRecipes: true,
    canShareContent: true,
    canAccessAdminPanel: true,
    canManageUsers: true,
    canViewAuditLogs: true,
    canExportData: true,
    maxQueriesPerDay: Infinity,
    maxRecipes: Infinity,
  },
  power_user: {
    canUseAI: true,
    canCreateRecipes: true,
    canShareContent: true,
    canAccessAdminPanel: false,
    canManageUsers: false,
    canViewAuditLogs: false,
    canExportData: true,
    maxQueriesPerDay: 1000,
    maxRecipes: 100,
  },
  standard_user: {
    canUseAI: true,
    canCreateRecipes: true,
    canShareContent: true,
    canAccessAdminPanel: false,
    canManageUsers: false,
    canViewAuditLogs: false,
    canExportData: false,
    maxQueriesPerDay: 100,
    maxRecipes: 20,
  },
  guest: {
    canUseAI: true,
    canCreateRecipes: false,
    canShareContent: false,
    canAccessAdminPanel: false,
    canManageUsers: false,
    canViewAuditLogs: false,
    canExportData: false,
    maxQueriesPerDay: 10,
    maxRecipes: 0,
  },
};

// ============================================================================
// Enterprise Auth Service
// ============================================================================

export class EnterpriseAuthService {
  private static instance: EnterpriseAuthService;
  private currentUser: EnterpriseUser | null = null;
  private currentSession: SessionInfo | null = null;
  private organization: Organization | null = null;
  private auditLogs: AuditLogEntry[] = [];
  private users: Map<string, EnterpriseUser> = new Map();
  private sessions: Map<string, SessionInfo> = new Map();
  private isInitialized: boolean = false;
  private failedLoginAttempts: Map<string, { count: number; lockedUntil?: Date }> = new Map();
  private readonly MAX_FAILED_ATTEMPTS = 5;
  private readonly LOCKOUT_DURATION_MS = 30 * 60 * 1000; // 30 minutes

  private constructor() {}

  static getInstance(): EnterpriseAuthService {
    if (!EnterpriseAuthService.instance) {
      EnterpriseAuthService.instance = new EnterpriseAuthService();
    }
    return EnterpriseAuthService.instance;
  }

  // ============================================================================
  // Initialization
  // ============================================================================

  /**
   * Initialize enterprise auth
   */
  async initialize(orgConfig?: Partial<Organization>): Promise<boolean> {
    try {
      // Load organization settings
      if (orgConfig) {
        this.organization = {
          id: orgConfig.id || `org-${Date.now()}`,
          name: orgConfig.name || 'Default Organization',
          domain: orgConfig.domain || 'localhost',
          ssoConfig: orgConfig.ssoConfig,
          settings: {
            requireMfa: false,
            sessionTimeoutMinutes: 480, // 8 hours
            maxUsers: 100,
            allowedProviders: ['local', 'azure_ad'],
            dataRetentionDays: 365,
            auditLogRetentionDays: 2555, // 7 years
            enableSharing: true,
            enableExternalAI: true,
            ...orgConfig.settings,
          },
          createdAt: new Date(),
        };
      }

      this.isInitialized = true;
      return true;
    } catch (error) {
      notificationManager.error('Failed to initialize enterprise auth: ' + error);
      return false;
    }
  }

  // ============================================================================
  // Authentication
  // ============================================================================

  /**
   * Sign in with SSO
   */
  async signInWithSSO(provider: AuthProvider): Promise<boolean> {
    this.logAudit('AUTH_SSO_INITIATED', 'authentication', { provider }, true);

    try {
      // In a real implementation, this would redirect to the SSO provider
      // For now, simulate successful authentication
      
      const mockUser: EnterpriseUser = {
        id: `user-${Date.now()}`,
        email: `user@${this.organization?.domain || 'example.com'}`,
        displayName: 'SSO User',
        role: 'standard_user',
        permissions: ROLE_PERMISSIONS.standard_user,
        provider,
        createdAt: new Date(),
        isActive: true,
        mfaEnabled: false,
      };

      await this.completeSignIn(mockUser);
      return true;
    } catch (error) {
      this.logAudit('AUTH_SSO_FAILED', 'authentication', { provider, error }, false);
      return false;
    }
  }

  /**
   * Sign in with local credentials
   */
  async signInWithLocal(email: string, password: string): Promise<boolean> {
    // Validate credentials (mock implementation)
    const user = this.users.get(email);
    
    if (!user || !user.isActive) {
      this.logAudit('AUTH_LOCAL_FAILED', 'authentication', { email, reason: 'invalid_credentials' }, false);
      notificationManager.error('Invalid credentials');
      return false;
    }

    // Check password (in real implementation, use bcrypt)
    if (password !== 'mock_password') {
      this.logAudit('AUTH_LOCAL_FAILED', 'authentication', { email, reason: 'wrong_password' }, false);
      notificationManager.error('Invalid credentials');
      return false;
    }

    await this.completeSignIn(user);
    return true;
  }

  private async completeSignIn(user: EnterpriseUser): Promise<void> {
    // Create session
    const sessionTimeout = this.organization?.settings.sessionTimeoutMinutes || 480;
    const session: SessionInfo = {
      sessionId: `session-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
      userId: user.id,
      startedAt: new Date(),
      expiresAt: new Date(Date.now() + sessionTimeout * 60 * 1000),
      isValid: true,
    };

    // Update user
    user.lastLoginAt = new Date();
    user.sessionExpiry = session.expiresAt;

    // Store
    this.currentUser = user;
    this.currentSession = session;
    this.sessions.set(session.sessionId, session);
    this.users.set(user.id, user);

    // Log success
    this.logAudit('AUTH_SUCCESS', 'authentication', { 
      userId: user.id, 
      email: user.email,
      role: user.role,
    }, true);

    notificationManager.success(`Welcome, ${user.displayName}!`);
  }

  /**
   * Sign out
   */
  async signOut(): Promise<void> {
    if (this.currentSession) {
      this.currentSession.isValid = false;
      this.sessions.delete(this.currentSession.sessionId);
      
      this.logAudit('AUTH_SIGNOUT', 'authentication', { 
        userId: this.currentUser?.id,
        sessionId: this.currentSession.sessionId,
      }, true);
    }

    this.currentUser = null;
    this.currentSession = null;
    notificationManager.info('Signed out successfully');
  }

  // ============================================================================
  // User Management
  // ============================================================================

  /**
   * Create a new user
   */
  async createUser(
    email: string,
    displayName: string,
    role: UserRole,
    department?: string
  ): Promise<EnterpriseUser | null> {
    if (!this.currentUser?.permissions.canManageUsers) {
      notificationManager.error('Insufficient permissions');
      return null;
    }

    // Check user limit
    if (this.users.size >= (this.organization?.settings.maxUsers || 100)) {
      notificationManager.error('Maximum user limit reached');
      return null;
    }

    const user: EnterpriseUser = {
      id: `user-${Date.now()}`,
      email,
      displayName,
      role,
      department,
      permissions: ROLE_PERMISSIONS[role],
      provider: 'local',
      createdAt: new Date(),
      isActive: true,
      mfaEnabled: false,
    };

    this.users.set(user.id, user);
    this.users.set(email, user); // Also index by email

    this.logAudit('USER_CREATED', 'user_management', { 
      targetUserId: user.id,
      email,
      role,
    }, true);

    notificationManager.success(`User ${email} created successfully`);
    return user;
  }

  /**
   * Update user
   */
  async updateUser(
    userId: string,
    updates: Partial<EnterpriseUser>
  ): Promise<EnterpriseUser | null> {
    if (!this.currentUser?.permissions.canManageUsers) {
      notificationManager.error('Insufficient permissions');
      return null;
    }

    const user = this.users.get(userId);
    if (!user) return null;

    // Update role permissions if role changed
    if (updates.role && updates.role !== user.role) {
      updates.permissions = ROLE_PERMISSIONS[updates.role];
    }

    const updatedUser = { ...user, ...updates, id: user.id };
    this.users.set(userId, updatedUser);

    this.logAudit('USER_UPDATED', 'user_management', { 
      targetUserId: userId,
      updates: Object.keys(updates),
    }, true);

    return updatedUser;
  }

  /**
   * Deactivate user
   */
  async deactivateUser(userId: string): Promise<boolean> {
    if (!this.currentUser?.permissions.canManageUsers) {
      notificationManager.error('Insufficient permissions');
      return false;
    }

    const user = this.users.get(userId);
    if (!user) return false;

    user.isActive = false;
    this.users.set(userId, user);

    // Invalidate all sessions
    for (const [sessionId, session] of this.sessions) {
      if (session.userId === userId) {
        session.isValid = false;
      }
    }

    this.logAudit('USER_DEACTIVATED', 'user_management', { targetUserId: userId }, true);
    logger.info('User deactivated', { targetUserId: userId, deactivatedBy: this.currentUser?.id });
    notificationManager.success(`User ${user.email} deactivated`);
    return true;
  }

  /**
   * Change user password
   */
  async changePassword(
    userId: string,
    currentPassword: string,
    newPassword: string
  ): Promise<boolean> {
    const user = this.users.get(userId);
    if (!user) {
      throw new PermissionError('User not found');
    }

    // Verify current password
    if (!this.verifyPassword(currentPassword, user)) {
      this.logAudit('PASSWORD_CHANGE_FAILED', 'user_management', { userId, reason: 'wrong_password' }, false);
      throw new PermissionError('Current password is incorrect');
    }

    // Validate password strength
    if (!this.isPasswordStrong(newPassword)) {
      throw new PermissionError(
        'New password must be at least 8 characters with uppercase, lowercase, number, and special character'
      );
    }

    // Update password (in production, hash the password)
    user.passwordHash = this.hashPassword(newPassword);
    this.users.set(userId, user);

    this.logAudit('PASSWORD_CHANGED', 'user_management', { userId }, true);
    logger.info('Password changed', { userId });
    notificationManager.success('Password changed successfully');
    return true;
  }

  // Private helper methods
  private recordFailedAttempt(email: string): void {
    const current = this.failedLoginAttempts.get(email) || { count: 0 };
    current.count++;
    
    if (current.count >= this.MAX_FAILED_ATTEMPTS) {
      current.lockedUntil = new Date(Date.now() + this.LOCKOUT_DURATION_MS);
      logger.warn('Account locked due to failed attempts', { email, attempts: current.count });
    }
    
    this.failedLoginAttempts.set(email, current);
  }

  /**
   * Verify password using constant-time comparison to prevent timing attacks
   */
  private verifyPassword(password: string, user: EnterpriseUser): boolean {
    if (!user.passwordHash) return false;
    
    const hashedInput = this.hashPassword(password);
    return this.constantTimeCompare(hashedInput, user.passwordHash);
  }

  /**
   * Constant-time string comparison to prevent timing attacks
   */
  private constantTimeCompare(a: string, b: string): boolean {
    if (a.length !== b.length) return false;
    
    let result = 0;
    for (let i = 0; i < a.length; i++) {
      result |= a.charCodeAt(i) ^ b.charCodeAt(i);
    }
    return result === 0;
  }

  private hashPassword(password: string): string {
    // In production, use bcrypt
    // This is a simple hash for demo
    let hash = 0;
    for (let i = 0; i < password.length; i++) {
      const char = password.charCodeAt(i);
      hash = ((hash << 5) - hash) + char;
      hash = hash & hash;
    }
    return hash.toString(16);
  }

  private isPasswordStrong(password: string): boolean {
    const minLength = password.length >= 8;
    const hasUpper = /[A-Z]/.test(password);
    const hasLower = /[a-z]/.test(password);
    const hasNumber = /[0-9]/.test(password);
    const hasSpecial = /[!@#$%^&*]/.test(password);
    return minLength && hasUpper && hasLower && hasNumber && hasSpecial;
  }

  /**
   * Get all users
   */
  getAllUsers(): EnterpriseUser[] {
    if (!this.currentUser?.permissions.canManageUsers) {
      return [];
    }
    return Array.from(this.users.values()).filter((user, index, self) => 
      self.findIndex(u => u.id === user.id) === index
    );
  }

  // ============================================================================
  // Permission Checks
  // ============================================================================

  /**
   * Check if user has permission
   */
  hasPermission(permission: keyof UserPermissions): boolean {
    return this.currentUser?.permissions[permission] === true;
  }

  /**
   * Check if user can perform action
   */
  canPerformAction(action: string): boolean {
    if (!this.currentUser) return false;

    const permissionMap: Record<string, keyof UserPermissions> = {
      'use_ai': 'canUseAI',
      'create_recipe': 'canCreateRecipes',
      'share': 'canShareContent',
      'access_admin': 'canAccessAdminPanel',
      'manage_users': 'canManageUsers',
      'view_audit': 'canViewAuditLogs',
      'export': 'canExportData',
    };

    const permission = permissionMap[action];
    return permission ? this.hasPermission(permission) : false;
  }

  // ============================================================================
  // Session Management
  // ============================================================================

  /**
   * Validate current session
   */
  isSessionValid(): boolean {
    if (!this.currentSession) return false;
    
    const now = new Date();
    return this.currentSession.isValid && this.currentSession.expiresAt > now;
  }

  /**
   * Refresh session
   */
  async refreshSession(): Promise<boolean> {
    if (!this.currentSession || !this.isSessionValid()) {
      return false;
    }

    const sessionTimeout = this.organization?.settings.sessionTimeoutMinutes || 480;
    this.currentSession.expiresAt = new Date(Date.now() + sessionTimeout * 60 * 1000);
    
    if (this.currentUser) {
      this.currentUser.sessionExpiry = this.currentSession.expiresAt;
    }

    return true;
  }

  /**
   * Get active sessions
   */
  getActiveSessions(): SessionInfo[] {
    if (!this.currentUser?.permissions.canManageUsers) {
      return this.currentSession ? [this.currentSession] : [];
    }

    const now = new Date();
    return Array.from(this.sessions.values()).filter(s => s.isValid && s.expiresAt > now);
  }

  /**
   * Revoke session
   */
  async revokeSession(sessionId: string): Promise<boolean> {
    if (!this.currentUser?.permissions.canManageUsers) {
      notificationManager.error('Insufficient permissions');
      return false;
    }

    const session = this.sessions.get(sessionId);
    if (!session) return false;

    session.isValid = false;
    this.sessions.delete(sessionId);

    this.logAudit('SESSION_REVOKED', 'session_management', { sessionId }, true);
    return true;
  }

  // ============================================================================
  // Audit Logging
  // ============================================================================

  /**
   * Log audit event
   */
  private logAudit(
    action: string,
    resource: string,
    details: Record<string, any>,
    success: boolean,
    errorMessage?: string
  ): void {
    const entry: AuditLogEntry = {
      id: `audit-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
      timestamp: new Date(),
      userId: this.currentUser?.id || 'anonymous',
      action,
      resource,
      details,
      success,
      errorMessage,
    };

    this.auditLogs.push(entry);

    // Trim old logs
    const maxLogs = 10000;
    if (this.auditLogs.length > maxLogs) {
      this.auditLogs = this.auditLogs.slice(-maxLogs);
    }
  }

  /**
   * Get audit logs
   */
  getAuditLogs(
    startDate?: Date,
    endDate?: Date,
    userId?: string,
    action?: string
  ): AuditLogEntry[] {
    if (!this.currentUser?.permissions.canViewAuditLogs) {
      notificationManager.error('Insufficient permissions');
      return [];
    }

    let logs = [...this.auditLogs];

    if (startDate) {
      logs = logs.filter(l => l.timestamp >= startDate);
    }
    if (endDate) {
      logs = logs.filter(l => l.timestamp <= endDate);
    }
    if (userId) {
      logs = logs.filter(l => l.userId === userId);
    }
    if (action) {
      logs = logs.filter(l => l.action === action);
    }

    return logs.sort((a, b) => b.timestamp.getTime() - a.timestamp.getTime());
  }

  /**
   * Export audit logs
   */
  exportAuditLogs(format: 'json' | 'csv' = 'json'): string {
    if (!this.currentUser?.permissions.canViewAuditLogs) {
      throw new Error('Insufficient permissions');
    }

    if (format === 'csv') {
      const headers = 'Timestamp,User ID,Action,Resource,Success,Details\n';
      const rows = this.auditLogs.map(l => 
        `"${l.timestamp.toISOString()}","${l.userId}","${l.action}","${l.resource}",${l.success},"${JSON.stringify(l.details).replace(/"/g, '""')}"`
      ).join('\n');
      return headers + rows;
    }

    return JSON.stringify(this.auditLogs, null, 2);
  }

  // ============================================================================
  // Getters
  // ============================================================================

  getCurrentUser(): EnterpriseUser | null {
    return this.currentUser;
  }

  getCurrentSession(): SessionInfo | null {
    return this.currentSession;
  }

  getOrganization(): Organization | null {
    return this.organization;
  }

  isAuthenticated(): boolean {
    return this.currentUser !== null && this.isSessionValid();
  }

  /**
   * Get access token for API authentication
   * Returns session ID as token for enterprise users
   */
  getAccessToken(): string | null {
    if (!this.isSessionValid() || !this.currentSession) {
      return null;
    }
    return this.currentSession.sessionId;
  }
}

// Export singleton instance
export const enterpriseAuth = EnterpriseAuthService.getInstance();
export default enterpriseAuth;
