/**
 * Unit Tests for EnterpriseAuthService
 * Tests authentication, session management, and security
 */

import { EnterpriseAuthService, UserRole, UserPermissions } from '../enterpriseAuth';

// Mock notificationManager
jest.mock('../../utils/notificationManager', () => ({
  notificationManager: {
    success: jest.fn(),
    error: jest.fn(),
    info: jest.fn(),
    warning: jest.fn()
  }
}));

// Mock logger
jest.mock('../../utils/logger', () => ({
  logger: {
    info: jest.fn(),
    error: jest.fn(),
    warn: jest.fn(),
    debug: jest.fn()
  }
}));

describe('EnterpriseAuthService', () => {
  let authService: EnterpriseAuthService;

  beforeEach(() => {
    // Reset singleton instance for each test
    (EnterpriseAuthService as any).instance = undefined;
    authService = EnterpriseAuthService.getInstance();
  });

  describe('initialization', () => {
    it('should be a singleton', () => {
      const instance1 = EnterpriseAuthService.getInstance();
      const instance2 = EnterpriseAuthService.getInstance();
      expect(instance1).toBe(instance2);
    });

    it('should initialize with default organization settings', async () => {
      const result = await authService.initialize();
      expect(result).toBe(true);
      
      const org = authService.getOrganization();
      expect(org).toBeDefined();
      expect(org?.settings.sessionTimeoutMinutes).toBe(480);
      expect(org?.settings.maxUsers).toBe(100);
    });

    it('should initialize with custom organization settings', async () => {
      const customConfig = {
        name: 'Test Org',
        domain: 'test.com',
        settings: {
          maxUsers: 50,
          requireMfa: true,
          sessionTimeoutMinutes: 240,
          allowedProviders: ['local' as const],
          dataRetentionDays: 180,
          auditLogRetentionDays: 365,
          enableSharing: true,
          enableExternalAI: false
        }
      };

      await authService.initialize(customConfig);
      const org = authService.getOrganization();
      
      expect(org?.name).toBe('Test Org');
      expect(org?.domain).toBe('test.com');
      expect(org?.settings.maxUsers).toBe(50);
      expect(org?.settings.requireMfa).toBe(true);
    });
  });

  describe('authentication', () => {
    beforeEach(async () => {
      await authService.initialize();
    });

    it('should sign in with SSO successfully', async () => {
      const result = await authService.signInWithSSO('azure_ad');
      expect(result).toBe(true);
      expect(authService.isAuthenticated()).toBe(true);
      
      const user = authService.getCurrentUser();
      expect(user).toBeDefined();
      expect(user?.provider).toBe('azure_ad');
      expect(user?.role).toBe('standard_user');
    });

    it('should create a session on sign in', async () => {
      await authService.signInWithSSO('local');
      
      const session = authService.getCurrentSession();
      expect(session).toBeDefined();
      expect(session?.isValid).toBe(true);
      expect(session?.expiresAt).toBeDefined();
    });

    it('should sign out successfully', async () => {
      await authService.signInWithSSO('local');
      expect(authService.isAuthenticated()).toBe(true);
      
      await authService.signOut();
      expect(authService.isAuthenticated()).toBe(false);
      expect(authService.getCurrentUser()).toBeNull();
    });

    it('should invalidate session on sign out', async () => {
      await authService.signInWithSSO('local');
      const session = authService.getCurrentSession();
      const sessionId = session?.sessionId;
      
      await authService.signOut();
      
      // Session should be invalidated
      const sessions = authService.getActiveSessions();
      expect(sessions.find(s => s.sessionId === sessionId)).toBeUndefined();
    });
  });

  describe('permission checks', () => {
    beforeEach(async () => {
      await authService.initialize();
    });

    it('should return false for permissions when not authenticated', () => {
      expect(authService.hasPermission('canUseAI')).toBe(false);
      expect(authService.canPerformAction('use_ai')).toBe(false);
    });

    it('should return correct permissions for standard user', async () => {
      await authService.signInWithSSO('local');
      
      expect(authService.hasPermission('canUseAI')).toBe(true);
      expect(authService.hasPermission('canCreateRecipes')).toBe(true);
      expect(authService.hasPermission('canShareContent')).toBe(true);
      expect(authService.hasPermission('canAccessAdminPanel')).toBe(false);
      expect(authService.hasPermission('canManageUsers')).toBe(false);
    });

    it('should check action permissions correctly', async () => {
      await authService.signInWithSSO('local');
      
      expect(authService.canPerformAction('use_ai')).toBe(true);
      expect(authService.canPerformAction('create_recipe')).toBe(true);
      expect(authService.canPerformAction('manage_users')).toBe(false);
    });
  });

  describe('session management', () => {
    beforeEach(async () => {
      await authService.initialize();
    });

    it('should validate active session', async () => {
      await authService.signInWithSSO('local');
      expect(authService.isSessionValid()).toBe(true);
    });

    it('should return false for session validation when not signed in', () => {
      expect(authService.isSessionValid()).toBe(false);
    });

    it('should get access token when authenticated', async () => {
      await authService.signInWithSSO('local');
      const token = authService.getAccessToken();
      expect(token).toBeDefined();
      expect(typeof token).toBe('string');
    });

    it('should return null for access token when not authenticated', () => {
      const token = authService.getAccessToken();
      expect(token).toBeNull();
    });

    it('should refresh session expiry', async () => {
      await authService.signInWithSSO('local');
      const sessionBefore = authService.getCurrentSession();
      const expiresBefore = sessionBefore?.expiresAt?.getTime();
      
      // Small delay to ensure time difference
      await new Promise(resolve => setTimeout(resolve, 10));
      
      await authService.refreshSession();
      const sessionAfter = authService.getCurrentSession();
      const expiresAfter = sessionAfter?.expiresAt?.getTime();
      
      expect(expiresAfter).toBeGreaterThan(expiresBefore || 0);
    });
  });

  describe('user management', () => {
    beforeEach(async () => {
      await authService.initialize();
      // Sign in as admin for user management tests
      await authService.signInWithSSO('local');
      // Promote to admin for testing
      const user = authService.getCurrentUser();
      if (user) {
        await authService.updateUser(user.id, { role: 'admin' });
      }
    });

    it('should create a new user', async () => {
      const newUser = await authService.createUser(
        'test@example.com',
        'Test User',
        'standard_user'
      );
      
      expect(newUser).toBeDefined();
      expect(newUser?.email).toBe('test@example.com');
      expect(newUser?.displayName).toBe('Test User');
      expect(newUser?.role).toBe('standard_user');
      expect(newUser?.isActive).toBe(true);
    });

    it('should update user role', async () => {
      const newUser = await authService.createUser(
        'update@example.com',
        'Update User',
        'standard_user'
      );
      
      expect(newUser).toBeDefined();
      
      const updated = await authService.updateUser(newUser!.id, {
        role: 'power_user'
      });
      
      expect(updated?.role).toBe('power_user');
      expect(updated?.permissions.maxQueriesPerDay).toBe(1000);
    });

    it('should deactivate user', async () => {
      const newUser = await authService.createUser(
        'deactivate@example.com',
        'Deactivate User',
        'standard_user'
      );
      
      expect(newUser).toBeDefined();
      
      const result = await authService.deactivateUser(newUser!.id);
      expect(result).toBe(true);
    });

    it('should get all users when admin', async () => {
      await authService.createUser('user1@example.com', 'User 1', 'standard_user');
      await authService.createUser('user2@example.com', 'User 2', 'standard_user');
      
      const users = authService.getAllUsers();
      expect(users.length).toBeGreaterThanOrEqual(3); // Admin + 2 users
    });
  });

  describe('password security', () => {
    it('should use constant-time comparison', () => {
      // Access private method for testing
      const compare = (authService as any).constantTimeCompare.bind(authService);
      
      expect(compare('abc', 'abc')).toBe(true);
      expect(compare('abc', 'abd')).toBe(false);
      expect(compare('abc', 'abcd')).toBe(false);
      expect(compare('', '')).toBe(true);
    });

    it('should validate password strength', () => {
      const isStrong = (authService as any).isPasswordStrong.bind(authService);
      
      expect(isStrong('weak')).toBe(false);
      expect(isStrong('password123')).toBe(false);
      expect(isStrong('Password123')).toBe(false);
      expect(isStrong('Password123!')).toBe(true);
      expect(isStrong('MyP@ssw0rd')).toBe(true);
    });
  });

  describe('audit logging', () => {
    beforeEach(async () => {
      await authService.initialize();
    });

    it('should log audit events on sign in', async () => {
      await authService.signInWithSSO('local');
      
      const logs = authService.getAuditLogs();
      expect(logs.length).toBeGreaterThan(0);
      
      const authLog = logs.find(l => l.action === 'AUTH_SUCCESS');
      expect(authLog).toBeDefined();
      expect(authLog?.success).toBe(true);
    });

    it('should export audit logs as JSON', async () => {
      await authService.signInWithSSO('local');
      
      const json = authService.exportAuditLogs('json');
      const parsed = JSON.parse(json);
      
      expect(Array.isArray(parsed)).toBe(true);
      expect(parsed.length).toBeGreaterThan(0);
    });

    it('should export audit logs as CSV', async () => {
      await authService.signInWithSSO('local');
      
      const csv = authService.exportAuditLogs('csv');
      expect(csv).toContain('Timestamp,User ID,Action,Resource,Success,Details');
    });
  });

  describe('role permissions', () => {
    const roleTests: Array<{ role: UserRole; expected: Partial<UserPermissions> }> = [
      {
        role: 'admin',
        expected: {
          canUseAI: true,
          canCreateRecipes: true,
          canShareContent: true,
          canAccessAdminPanel: true,
          canManageUsers: true,
          canViewAuditLogs: true,
          canExportData: true,
          maxQueriesPerDay: Infinity,
          maxRecipes: Infinity
        }
      },
      {
        role: 'power_user',
        expected: {
          canUseAI: true,
          canCreateRecipes: true,
          canShareContent: true,
          canAccessAdminPanel: false,
          canManageUsers: false,
          canExportData: true,
          maxQueriesPerDay: 1000,
          maxRecipes: 100
        }
      },
      {
        role: 'standard_user',
        expected: {
          canUseAI: true,
          canCreateRecipes: true,
          canShareContent: true,
          canAccessAdminPanel: false,
          canManageUsers: false,
          canExportData: false,
          maxQueriesPerDay: 100,
          maxRecipes: 20
        }
      },
      {
        role: 'guest',
        expected: {
          canUseAI: true,
          canCreateRecipes: false,
          canShareContent: false,
          canAccessAdminPanel: false,
          canManageUsers: false,
          canExportData: false,
          maxQueriesPerDay: 10,
          maxRecipes: 0
        }
      }
    ];

    roleTests.forEach(({ role, expected }) => {
      it(`should have correct permissions for ${role}`, async () => {
        await authService.signInWithSSO('local');
        const user = authService.getCurrentUser();
        
        if (user) {
          await authService.updateUser(user.id, { role });
          
          const updatedUser = authService.getCurrentUser();
          expect(updatedUser?.role).toBe(role);
          
          Object.entries(expected).forEach(([key, value]) => {
            expect(updatedUser?.permissions[key as keyof UserPermissions]).toBe(value);
          });
        }
      });
    });
  });
});