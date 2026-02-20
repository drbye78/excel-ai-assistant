// Empty State Component - Helpful empty states with clear CTAs
// Replaces blank screens with actionable guidance

import React from 'react';
import { Stack, Text, PrimaryButton, DefaultButton, Icon, Link } from '@fluentui/react';

export interface EmptyStateProps {
  /** Fluent UI icon name */
  icon: string;
  /** Main title text */
  title: string;
  /** Descriptive text explaining the empty state */
  description: string;
  /** Optional illustration/image URL */
  illustration?: string;
  /** Primary action button */
  primaryAction?: {
    text: string;
    icon?: string;
    onClick: () => void;
  };
  /** Secondary action button */
  secondaryAction?: {
    text: string;
    onClick: () => void;
  };
  /** Link text for alternative action */
  linkAction?: {
    text: string;
    href?: string;
    onClick?: () => void;
  };
  /** Compact mode for inline empty states */
  compact?: boolean;
  /** Custom styles override */
  styles?: React.CSSProperties;
}

/**
 * EmptyState Component
 * 
 * Usage examples:
 * ```tsx
 * // Full page empty state
 * <EmptyState
 *   icon="History"
 *   title="No conversations yet"
 *   description="Start chatting with the AI assistant to see your conversation history here."
 *   primaryAction={{ text: 'Start New Chat', icon: 'Chat', onClick: () => navigate('/chat') }}
 * />
 * 
 * // Compact inline state
 * <EmptyState
 *   icon="Search"
 *   title="No results found"
 *   description="Try adjusting your search criteria."
 *   compact
 * />
 * ```
 */
export const EmptyState: React.FC<EmptyStateProps> = ({
  icon,
  title,
  description,
  illustration,
  primaryAction,
  secondaryAction,
  linkAction,
  compact = false,
  styles: customStyles,
}) => {
  const iconSize = compact ? 32 : 64;
  const padding = compact ? 16 : 40;
  const gap = compact ? 8 : 15;

  return (
    <Stack
      horizontalAlign="center"
      tokens={{ childrenGap: gap, padding }}
      styles={{
        root: {
          width: '100%',
          ...customStyles,
        },
      }}
    >
      {illustration ? (
        <img
          src={illustration}
          alt=""
          width={compact ? 80 : 120}
          height={compact ? 80 : 120}
          style={{ marginBottom: compact ? 8 : 16 }}
        />
      ) : (
        <Icon
          iconName={icon}
          styles={{
            root: {
              fontSize: iconSize,
              color: '#605e5c',
              marginBottom: compact ? 4 : 8,
            },
          }}
        />
      )}

      <Stack tokens={{ childrenGap: compact ? 4 : 8 }} horizontalAlign="center">
        <Text
          variant={compact ? 'medium' : 'xLarge'}
          styles={{
            root: {
              fontWeight: 600,
              textAlign: 'center',
            },
          }}
        >
          {title}
        </Text>
        <Text
          styles={{
            root: {
              color: '#605e5c',
              maxWidth: compact ? 300 : 400,
              textAlign: 'center',
              fontSize: compact ? 12 : 14,
            },
          }}
        >
          {description}
        </Text>
      </Stack>

      {(primaryAction || secondaryAction) && (
        <Stack
          horizontal={!compact}
          vertical={compact}
          tokens={{ childrenGap: compact ? 8 : 12 }}
          horizontalAlign="center"
          styles={{ root: { marginTop: compact ? 8 : 16 } }}
        >
          {primaryAction && (
            <PrimaryButton
              iconProps={primaryAction.icon ? { iconName: primaryAction.icon } : undefined}
              onClick={primaryAction.onClick}
            >
              {primaryAction.text}
            </PrimaryButton>
          )}
          {secondaryAction && (
            <DefaultButton onClick={secondaryAction.onClick}>
              {secondaryAction.text}
            </DefaultButton>
          )}
        </Stack>
      )}

      {linkAction && (
        <Link
          onClick={linkAction.onClick}
          href={linkAction.href}
          target={linkAction.href?.startsWith('http') ? '_blank' : undefined}
          styles={{ root: { marginTop: compact ? 4 : 8 } }}
        >
          {linkAction.text}
        </Link>
      )}
    </Stack>
  );
};

// Pre-configured empty states for common scenarios
export const EmptyStates = {
  /**
   * No data available
   */
  NoData: (props: Omit<EmptyStateProps, 'icon' | 'title'>) => (
    <EmptyState
      icon="Page"
      title="No data available"
      description="There's nothing to show here yet."
      {...props}
    />
  ),

  /**
   * No search results
   */
  NoResults: (props: Omit<EmptyStateProps, 'icon' | 'title'>) => (
    <EmptyState
      icon="Search"
      title="No results found"
      description="We couldn't find anything matching your search. Try different keywords."
      {...props}
    />
  ),

  /**
   * No conversations/history
   */
  NoHistory: (props: Omit<EmptyStateProps, 'icon' | 'title'>) => (
    <EmptyState
      icon="History"
      title="No history yet"
      description="Your activity will appear here once you start using the features."
      {...props}
    />
  ),

  /**
   * No recipes saved
   */
  NoRecipes: (props: Omit<EmptyStateProps, 'icon' | 'title'>) => (
    <EmptyState
      icon="ClipboardList"
      title="No recipes yet"
      description="Save your favorite operations as recipes to reuse them quickly."
      {...props}
    />
  ),

  /**
   * Error state
   */
  Error: (props: Omit<EmptyStateProps, 'icon' | 'title'>) => (
    <EmptyState
      icon="Error"
      title="Something went wrong"
      description="We couldn't load the data. Please try again."
      {...props}
    />
  ),

  /**
   * Coming soon
   */
  ComingSoon: (props: Omit<EmptyStateProps, 'icon' | 'title'>) => (
    <EmptyState
      icon="ConstructionCone"
      title="Coming soon"
      description="This feature is under development. Check back later!"
      {...props}
    />
  ),
};

export default EmptyState;
