/**
 * Compliance Suite Service
 *
 * Provides data governance, compliance monitoring, and regulatory
 * adherence features for enterprise deployments.
 *
 * Features:
 * - GDPR compliance (data export, deletion)
 * - Data retention policies
 * - PII detection and redaction
 * - Compliance reporting
 * - Data classification
 *
 * @module services/complianceService
 */

import { notificationManager } from '../utils/notificationManager';
import { logger } from '../utils/logger';

export type DataClassification = 'public' | 'internal' | 'confidential' | 'restricted';
export type ComplianceStatus = 'compliant' | 'warning' | 'violation';

export interface CompliancePolicy {
  id: string;
  name: string;
  type: 'gdpr' | 'hipaa' | 'sox' | 'custom';
  enabled: boolean;
  settings: Record<string, any>;
}

export interface DataRetentionRule {
  id: string;
  dataType: string;
  retentionDays: number;
  autoDelete: boolean;
}

export interface ComplianceReport {
  generatedAt: Date;
  overallStatus: ComplianceStatus;
  policies: PolicyCompliance[];
  violations: ComplianceViolation[];
  recommendations: string[];
}

export interface PolicyCompliance {
  policyId: string;
  policyName: string;
  status: ComplianceStatus;
  details: string;
}

export interface ComplianceViolation {
  id: string;
  policyId: string;
  severity: 'low' | 'medium' | 'high' | 'critical';
  message: string;
  resource: string;
  detectedAt: Date;
}

export interface PIIEntity {
  type: 'email' | 'phone' | 'ssn' | 'credit_card' | 'name' | 'address';
  value: string;
  position: number;
  confidence: number;
}

// GDPR-specific interfaces
export interface UserData {
  id: string;
  email: string;
  name: string;
  consent?: boolean;
  consentRecord?: boolean;
  consentTimestamp?: Date;
  contractual?: boolean;
  legal?: boolean;
  retentionPolicy?: string;
  processingPurpose?: string;
  createdAt: Date;
}

export interface GDPRCheckResult {
  compliant: boolean;
  violations: GDPRViolation[];
  recommendations: string[];
  dataInventory: DataInventoryItem[];
}

export interface GDPRViolation {
  article: string;
  severity: 'critical' | 'high' | 'medium' | 'low';
  description: string;
  remediation: string;
}

export interface DataInventoryItem {
  category: string;
  type: 'personal' | 'sensitive' | 'financial' | 'other';
  retention: string;
  lawfulBasis?: string;
}

export interface GDPRDataExport {
  userId: string;
  exportDate: Date;
  data: {
    profile: any;
    conversations: any[];
    analytics: any;
    settings: any;
    consentHistory: any[];
    auditLogs: any[];
  };
  format: 'json' | 'csv';
  checksum: string;
}

export interface DeleteOptions {
  keepAnonymized?: boolean;
  legalHold?: boolean;
}

export interface DeletionResult {
  userId: string;
  deletedAt: Date;
  itemsDeleted: string[];
  itemsRetained: string[];
  retentionReasons: Record<string, string>;
}

export class ComplianceService {
  private static instance: ComplianceService;
  private policies: Map<string, CompliancePolicy> = new Map();
  private violations: ComplianceViolation[] = [];
  private retentionRules: Map<string, DataRetentionRule> = new Map();

  private constructor() {
    this.initializeDefaultPolicies();
  }

  static getInstance(): ComplianceService {
    if (!ComplianceService.instance) {
      ComplianceService.instance = new ComplianceService();
    }
    return ComplianceService.instance;
  }

  private initializeDefaultPolicies(): void {
    this.policies.set('gdpr', {
      id: 'gdpr',
      name: 'GDPR Compliance',
      type: 'gdpr',
      enabled: true,
      settings: {
        allowDataExport: true,
        requireConsent: true,
        dataRetentionDays: 365,
      },
    });

    this.retentionRules.set('conversation', {
      id: 'conversation',
      dataType: 'conversation_history',
      retentionDays: 365,
      autoDelete: false,
    });

    this.retentionRules.set('audit', {
      id: 'audit',
      dataType: 'audit_logs',
      retentionDays: 2555, // 7 years
      autoDelete: false,
    });
  }

  /**
   * Scan text for PII
   */
  detectPII(text: string): PIIEntity[] {
    const entities: PIIEntity[] = [];
    
    // Email pattern
    const emailRegex = /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b/g;
    let match;
    while ((match = emailRegex.exec(text)) !== null) {
      entities.push({
        type: 'email',
        value: match[0],
        position: match.index,
        confidence: 0.95,
      });
    }

    // Phone pattern
    const phoneRegex = /\b\d{3}[-.]?\d{3}[-.]?\d{4}\b/g;
    while ((match = phoneRegex.exec(text)) !== null) {
      entities.push({
        type: 'phone',
        value: match[0],
        position: match.index,
        confidence: 0.9,
      });
    }

    // SSN pattern
    const ssnRegex = /\b\d{3}-\d{2}-\d{4}\b/g;
    while ((match = ssnRegex.exec(text)) !== null) {
      entities.push({
        type: 'ssn',
        value: match[0],
        position: match.index,
        confidence: 0.85,
      });
    }

    return entities;
  }

  /**
   * Redact PII from text
   */
  redactPII(text: string): string {
    const entities = this.detectPII(text);
    let redacted = text;
    
    // Sort by position descending to avoid offset issues
    entities.sort((a, b) => b.position - a.position);
    
    for (const entity of entities) {
      const redaction = `[${entity.type.toUpperCase()}_REDACTED]`;
      redacted = redacted.substring(0, entity.position) + 
                 redaction + 
                 redacted.substring(entity.position + entity.value.length);
    }
    
    return redacted;
  }

  /**
   * Generate compliance report
   */
  async generateReport(): Promise<ComplianceReport> {
    const violations: ComplianceViolation[] = [];
    const policies: PolicyCompliance[] = [];

    // Check each policy
    for (const [id, policy] of this.policies) {
      if (!policy.enabled) continue;

      const status = await this.checkPolicyCompliance(policy);
      policies.push({
        policyId: id,
        policyName: policy.name,
        status: status.status,
        details: status.details,
      });

      violations.push(...status.violations);
    }

    const overallStatus = violations.some(v => v.severity === 'critical' || v.severity === 'high')
      ? 'violation'
      : violations.length > 0
      ? 'warning'
      : 'compliant';

    return {
      generatedAt: new Date(),
      overallStatus,
      policies,
      violations,
      recommendations: this.generateRecommendations(violations),
    };
  }

  private async checkPolicyCompliance(policy: CompliancePolicy): Promise<{
    status: ComplianceStatus;
    details: string;
    violations: ComplianceViolation[];
  }> {
    const violations: ComplianceViolation[] = [];
    
    // Mock compliance checks
    if (policy.type === 'gdpr') {
      // Check data retention
      for (const [id, rule] of this.retentionRules) {
        if (rule.retentionDays > 2555) {
          violations.push({
            id: `viol-${Date.now()}`,
            policyId: policy.id,
            severity: 'medium',
            message: `Retention period for ${rule.dataType} exceeds recommended limit`,
            resource: rule.dataType,
            detectedAt: new Date(),
          });
        }
      }
    }

    return {
      status: violations.length === 0 ? 'compliant' : 'violation',
      details: `Checked ${policy.name}`,
      violations,
    };
  }

  private generateRecommendations(violations: ComplianceViolation[]): string[] {
    const recommendations: string[] = [];
    
    if (violations.some(v => v.severity === 'critical')) {
      recommendations.push('Address critical violations immediately');
    }
    if (violations.some(v => v.message.includes('retention'))) {
      recommendations.push('Review and update data retention policies');
    }
    if (violations.length === 0) {
      recommendations.push('Continue monitoring compliance status');
    }

    return recommendations;
  }

  /**
   * Perform comprehensive GDPR compliance audit
   */
  async performGDPRAudit(userData: UserData): Promise<GDPRCheckResult> {
    const violations: GDPRViolation[] = [];
    const recommendations: string[] = [];
    const dataInventory: DataInventoryItem[] = [];

    // Check 1: Lawful basis for processing (Article 6)
    if (!userData.consent && !userData.contractual && !userData.legal) {
      violations.push({
        article: 'Article 6',
        severity: 'critical',
        description: 'No lawful basis for processing personal data',
        remediation: 'Obtain explicit consent or establish another lawful basis'
      });
    }

    // Check 2: Data minimization (Article 5(1)(c))
    const unnecessaryFields = this.identifyUnnecessaryFields(userData);
    if (unnecessaryFields.length > 0) {
      violations.push({
        article: 'Article 5(1)(c)',
        severity: 'high',
        description: `Collecting unnecessary data: ${unnecessaryFields.join(', ')}`,
        remediation: 'Remove fields not essential for the specified purpose'
      });
    }

    // Check 3: Storage limitation (Article 5(1)(e))
    if (!userData.retentionPolicy) {
      violations.push({
        article: 'Article 5(1)(e)',
        severity: 'medium',
        description: 'No data retention policy defined',
        remediation: 'Define and implement data retention periods'
      });
    }

    // Check 4: Purpose limitation (Article 5(1)(b))
    if (!userData.processingPurpose) {
      violations.push({
        article: 'Article 5(1)(b)',
        severity: 'high',
        description: 'Processing purpose not specified',
        remediation: 'Document specific purposes for data processing'
      });
    }

    // Check 5: Consent records (Article 7)
    if (userData.consent) {
      if (!userData.consentRecord || !userData.consentTimestamp) {
        violations.push({
          article: 'Article 7',
          severity: 'critical',
          description: 'Consent obtained but no record maintained',
          remediation: 'Implement consent logging with timestamp and version'
        });
      }
    }

    // Build data inventory
    dataInventory.push(
      { category: 'Email', type: 'personal', retention: userData.retentionPolicy || 'Not defined', lawfulBasis: userData.consent ? 'Consent' : 'Legitimate Interest' },
      { category: 'Name', type: 'personal', retention: userData.retentionPolicy || 'Not defined', lawfulBasis: userData.consent ? 'Consent' : 'Contract' },
      { category: 'Usage Data', type: 'other', retention: userData.retentionPolicy || '90 days', lawfulBasis: 'Legitimate Interest' },
      { category: 'Conversations', type: 'personal', retention: userData.retentionPolicy || '365 days', lawfulBasis: 'Consent' }
    );

    // Add recommendations
    recommendations.push(
      'Implement data protection impact assessment (DPIA)',
      'Appoint a Data Protection Officer if processing is core business',
      'Implement privacy by design principles',
      'Establish data breach response procedure',
      'Create data subject request handling process'
    );

    logger.info('GDPR audit completed', { 
      userId: userData.id, 
      violations: violations.length,
      compliant: violations.length === 0 
    });

    return {
      compliant: violations.length === 0,
      violations,
      recommendations,
      dataInventory
    };
  }

  private identifyUnnecessaryFields(userData: UserData): string[] {
    const essential = ['id', 'email', 'name', 'role', 'createdAt'];
    const collected = Object.keys(userData);
    return collected.filter(field => !essential.includes(field) && !field.startsWith('consent'));
  }

  /**
   * Export user data (GDPR Article 20 - Right to data portability)
   */
  async exportUserData(userId: string): Promise<GDPRDataExport> {
    logger.info('Starting GDPR data export', { userId });

    try {
      // Collect all data related to the user
      const exportData: GDPRDataExport = {
        userId,
        exportDate: new Date(),
        data: {
          profile: await this.getUserProfile(userId),
          conversations: await this.getUserConversations(userId),
          analytics: await this.getUserAnalytics(userId),
          settings: await this.getUserSettings(userId),
          consentHistory: await this.getConsentHistory(userId),
          auditLogs: await this.getUserAuditLogs(userId)
        },
        format: 'json',
        checksum: ''
      };

      // Generate checksum for data integrity
      exportData.checksum = await this.generateChecksum(exportData.data);

      logger.info('GDPR data export completed', { 
        userId, 
        size: JSON.stringify(exportData).length,
        checksum: exportData.checksum
      });

      notificationManager.success(`Data exported for user ${userId}`);
      return exportData;
    } catch (error) {
      logger.error('Failed to export user data', { userId }, error as Error);
      throw new Error(`Failed to export data: ${(error as Error).message}`);
    }
  }

  /**
   * Delete user data (GDPR Article 17 - Right to be forgotten)
   */
  async deleteUserData(userId: string, options: DeleteOptions = {}): Promise<DeletionResult> {
    logger.info('Starting GDPR data deletion', { userId, options });

    const result: DeletionResult = {
      userId,
      deletedAt: new Date(),
      itemsDeleted: [],
      itemsRetained: [],
      retentionReasons: {}
    };

    try {
      // Check for legal obligations that prevent deletion
      const legalHold = await this.checkLegalHold(userId);
      if (legalHold.hasHold) {
        result.itemsRetained.push('account');
        result.retentionReasons['account'] = legalHold.reason;
        logger.warn('Legal hold prevents deletion', { userId, reason: legalHold.reason });
      } else {
        // Delete user profile
        await this.deleteUserProfile(userId);
        result.itemsDeleted.push('profile');
      }

      // Handle conversations based on options
      if (options.keepAnonymized) {
        await this.anonymizeConversations(userId);
        result.itemsRetained.push('conversations (anonymized)');
        result.retentionReasons['conversations'] = 'Anonymized for analytics retention';
      } else {
        await this.deleteConversations(userId);
        result.itemsDeleted.push('conversations');
      }

      // Delete analytics (or anonymize)
      await this.deleteUserAnalytics(userId);
      result.itemsDeleted.push('analytics');

      // Delete settings
      await this.deleteUserSettings(userId);
      result.itemsDeleted.push('settings');

      // Delete consent history
      await this.deleteConsentHistory(userId);
      result.itemsDeleted.push('consent_history');

      logger.info('GDPR data deletion completed', result);
      notificationManager.success(`All data deleted for user ${userId}`);
      return result;
    } catch (error) {
      logger.error('Failed to delete user data', { userId }, error as Error);
      throw new Error(`Failed to delete data: ${(error as Error).message}`);
    }
  }

  // Private helper methods for data operations
  private async getUserProfile(userId: string): Promise<any> {
    // Would fetch from user service
    return { userId, exported: true };
  }

  private async getUserConversations(userId: string): Promise<any[]> {
    // Would fetch from conversation storage
    return [];
  }

  private async getUserAnalytics(userId: string): Promise<any> {
    // Would fetch from analytics service
    return { userId, exported: true };
  }

  private async getUserSettings(userId: string): Promise<any> {
    // Would fetch from settings storage
    return {};
  }

  private async getConsentHistory(userId: string): Promise<any[]> {
    // Would fetch from consent records
    return [];
  }

  private async getUserAuditLogs(userId: string): Promise<any[]> {
    // Would fetch from audit log
    return [];
  }

  private async deleteUserProfile(userId: string): Promise<void> {
    // Would delete from user service
    logger.debug('User profile deleted', { userId });
  }

  private async deleteConversations(userId: string): Promise<void> {
    // Would delete from conversation storage
    logger.debug('User conversations deleted', { userId });
  }

  private async anonymizeConversations(userId: string): Promise<void> {
    // Would anonymize conversations for retention
    logger.debug('User conversations anonymized', { userId });
  }

  private async deleteUserAnalytics(userId: string): Promise<void> {
    // Would delete from analytics
    logger.debug('User analytics deleted', { userId });
  }

  private async deleteUserSettings(userId: string): Promise<void> {
    // Would delete settings
    logger.debug('User settings deleted', { userId });
  }

  private async deleteConsentHistory(userId: string): Promise<void> {
    // Would delete consent records
    logger.debug('Consent history deleted', { userId });
  }

  private async checkLegalHold(userId: string): Promise<{ hasHold: boolean; reason?: string }> {
    // Check if user data is under legal hold
    // This would query a legal hold database
    return { hasHold: false };
  }

  private async generateChecksum(data: any): Promise<string> {
    // Simple checksum for data integrity verification
    const str = JSON.stringify(data);
    let hash = 0;
    for (let i = 0; i < str.length; i++) {
      const char = str.charCodeAt(i);
      hash = ((hash << 5) - hash) + char;
      hash = hash & hash;
    }
    return hash.toString(16).padStart(8, '0');
  }

  /**
   * Get retention rules
   */
  getRetentionRules(): DataRetentionRule[] {
    return Array.from(this.retentionRules.values());
  }

  /**
   * Update retention rule
   */
  updateRetentionRule(ruleId: string, updates: Partial<DataRetentionRule>): boolean {
    const rule = this.retentionRules.get(ruleId);
    if (!rule) return false;
    
    Object.assign(rule, updates);
    this.retentionRules.set(ruleId, rule);
    return true;
  }
}

export const complianceService = ComplianceService.getInstance();
export default complianceService;
