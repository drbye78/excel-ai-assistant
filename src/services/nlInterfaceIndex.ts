// Natural Language Interface Services Index
// Central export point for all NL interface services

// Core parser exports
export { NaturalLanguageCommandParser } from './naturalLanguageCommandParser';
export type {
  ParsedCommand,
  NLContext,
  CommandIntent,
  CommandTarget,
  SupportedLocale,
  ConversationState
} from './naturalLanguageCommandParser';

// Smart suggestion exports
export { SmartSuggestionEngine } from './smartSuggestionEngine';
export type {
  SuggestionCategory,
  ContextualSuggestion
} from './smartSuggestionEngine';

// Error recovery exports
export { ErrorRecoveryEngine } from './errorRecoveryEngine';
export type {
  RecoverySuggestion,
  ErrorRecoveryResult
} from './errorRecoveryEngine';

// Clarification exports
export { ClarificationEngine } from './clarificationEngine';
export type {
  ClarificationRequest,
  ClarificationResponse,
  ClarificationOption,
  ClarificationType
} from './clarificationEngine';

// Named Range exports
export { NamedRangeService } from './namedRangeService';
export type {
  NamedRange,
  NamedRangeCreateOptions,
  NamedRangeUpdateOptions
} from './namedRangeService';

// Comment exports
export { CommentService } from './commentService';
export type {
  CellComment,
  CommentCreateOptions,
  CommentUpdateOptions
} from './commentService';

// Validation exports
export { ValidationService } from './validationService';
export type {
  ValidationRule,
  ValidationCheckResult
} from './validationService';

// Hyperlink exports
export { HyperlinkService } from './hyperlinkService';
export type {
  Hyperlink,
  HyperlinkCreateOptions
} from './hyperlinkService';

// Macro exports
export { MacroService } from './macroService';
export type {
  Macro,
  MacroRecording,
  MacroAction
} from './macroService';

// Convenience re-exports for direct use
export { default as nlParser } from './naturalLanguageCommandParser';
export { default as suggestionEngine } from './smartSuggestionEngine';
export { default as errorRecovery } from './errorRecoveryEngine';
export { default as clarificationEngine } from './clarificationEngine';
export { default as namedRangeService } from './namedRangeService';
export { default as commentService } from './commentService';
export { default as validationService } from './validationService';
export { default as hyperlinkService } from './hyperlinkService';
export { default as macroService } from './macroService';
