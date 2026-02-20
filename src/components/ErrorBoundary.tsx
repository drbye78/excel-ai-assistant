// Error Boundary Component - Production-ready error handling
// Prevents entire add-in crash when a component fails

import * as React from 'react';
import { Component, ErrorInfo, ReactNode } from 'react';
import { Stack, Text, PrimaryButton, DefaultButton, Icon } from '@fluentui/react';
import logger from '@/utils/logger';

interface Props {
  children: ReactNode;
  section?: string;
  fallback?: ReactNode;
  onReset?: () => void;
}

interface State {
  hasError: boolean;
  error: Error | null;
  errorInfo: ErrorInfo | null;
}

export class ErrorBoundary extends Component<Props, State> {
  public state: State = {
    hasError: false,
    error: null,
    errorInfo: null,
  };

  public static getDerivedStateFromError(error: Error): State {
    return { hasError: true, error, errorInfo: null };
  }

  public componentDidCatch(error: Error, errorInfo: ErrorInfo) {
    logger.error(`Error in ${this.props.section || 'component'}`, { componentStack: errorInfo.componentStack }, error);
    this.setState({ errorInfo });
    
    // Log to analytics/monitoring service
    this.logError(error, errorInfo);
  }

  private logError(error: Error, errorInfo: ErrorInfo) {
    // In production, send to error tracking service
    if (process.env.NODE_ENV === 'production') {
      // Example: Sentry, Application Insights, etc.
      // logErrorToService({ error, errorInfo, section: this.props.section });
    }
  }

  private handleReset = () => {
    this.setState({ hasError: false, error: null, errorInfo: null });
    this.props.onReset?.();
  };

  private handleReload = () => {
    window.location.reload();
  };

  private handleReportIssue = () => {
    const { error, errorInfo } = this.state;
    const issueBody = encodeURIComponent(
      `**Section:** ${this.props.section || 'Unknown'}\n\n` +
      `**Error:** ${error?.message}\n\n` +
      `**Stack Trace:**\n\`\`\`\n${error?.stack}\n\`\`\`\n\n` +
      `**Component Stack:**\n\`\`\`\n${errorInfo?.componentStack}\n\`\`\``
    );
    window.open(`https://github.com/your-repo/issues/new?body=${issueBody}`, '_blank');
  };

  public render() {
    if (this.state.hasError) {
      // Custom fallback UI
      if (this.props.fallback) {
        return this.props.fallback;
      }

      return (
        <Stack
          horizontalAlign="center"
          verticalAlign="center"
          tokens={{ childrenGap: 16, padding: 24 }}
          styles={{
            root: {
              minHeight: 300,
              backgroundColor: '#faf9f8',
              borderRadius: 4,
            },
          }}
        >
          <Icon
            iconName="ErrorBadge"
            styles={{
              root: {
                fontSize: 48,
                color: '#d13438',
              },
            }}
          />
          
          <Stack tokens={{ childrenGap: 8 }} horizontalAlign="center">
            <Text variant="xLarge" styles={{ root: { fontWeight: 600 } }}>
              Something went wrong
            </Text>
            <Text styles={{ root: { color: '#605e5c', textAlign: 'center', maxWidth: 400 } }}>
              {this.props.section 
                ? `We encountered an error in the ${this.props.section} section.` 
                : 'We encountered an unexpected error.'}
            </Text>
            {process.env.NODE_ENV === 'development' && this.state.error && (
              <Text
                styles={{
                  root: {
                    color: '#d13438',
                    fontSize: 12,
                    fontFamily: 'monospace',
                    backgroundColor: '#fed9cc',
                    padding: 8,
                    borderRadius: 4,
                    maxWidth: 500,
                    overflow: 'auto',
                  },
                }}
              >
                {this.state.error.message}
              </Text>
            )}
          </Stack>

          <Stack horizontal tokens={{ childrenGap: 12 }}>
            <PrimaryButton
              iconProps={{ iconName: 'Refresh' }}
              onClick={this.handleReset}
            >
              Try Again
            </PrimaryButton>
            <DefaultButton
              iconProps={{ iconName: 'Refresh' }}
              onClick={this.handleReload}
            >
              Reload Add-in
            </DefaultButton>
          </Stack>

          <DefaultButton
            iconProps={{ iconName: 'ReportHacked' }}
            onClick={this.handleReportIssue}
            styles={{ root: { marginTop: 8 } }}
          >
            Report Issue
          </DefaultButton>
        </Stack>
      );
    }

    return this.props.children;
  }
}

// Hook for functional components to handle errors gracefully
export const useErrorHandler = () => {
  const [error, setError] = React.useState<Error | null>(null);

  const handleError = React.useCallback((err: Error) => {
    logger.error('Handled error', undefined, err);
    setError(err);
  }, []);

  const clearError = React.useCallback(() => {
    setError(null);
  }, []);

  return { error, handleError, clearError };
};

export default ErrorBoundary;
