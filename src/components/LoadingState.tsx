// Loading State Component - Skeleton loading patterns
// Provides consistent loading feedback across the app

import React from 'react';
import { Shimmer, ShimmerElementType, Stack, Spinner, SpinnerSize, Text } from '@fluentui/react';

export type LoadingType = 'card' | 'list' | 'detail' | 'form' | 'table' | 'text' | 'custom';

export interface LoadingStateProps {
  /** Type of loading pattern to display */
  type: LoadingType;
  /** Number of items to show (for list/table types) */
  count?: number;
  /** Custom shimmer elements */
  customElements?: Array<{ type: ShimmerElementType; height: number; width?: number }>;
  /** Loading message */
  message?: string;
  /** Show spinner instead of shimmer */
  useSpinner?: boolean;
  /** Custom height */
  height?: number;
}

/**
 * LoadingState Component
 * 
 * Usage examples:
 * ```tsx
 * // Card loading
 * <LoadingState type="card" />
 * 
 * // List with 5 items
 * <LoadingState type="list" count={5} />
 * 
 * // With spinner
 * <LoadingState type="detail" useSpinner message="Loading recipe..." />
 * ```
 */
export const LoadingState: React.FC<LoadingStateProps> = ({
  type,
  count = 3,
  customElements,
  message,
  useSpinner = false,
  height,
}) => {
  if (useSpinner) {
    return (
      <Stack
        horizontalAlign="center"
        verticalAlign="center"
        tokens={{ childrenGap: 12, padding: 24 }}
        styles={{ root: { height: height || 200 } }}
      >
        <Spinner size={SpinnerSize.large} />
        {message && <Text styles={{ root: { color: '#605e5c' } }}>{message}</Text>}
      </Stack>
    );
  }

  const renderShimmer = (elements: Array<{ type: ShimmerElementType; height: number; width?: number }>) => (
    <Shimmer
      shimmerElements={elements}
      styles={{
        root: {
          marginBottom: 8,
        },
      }}
    />
  );

  const patterns: Record<LoadingType, React.ReactNode> = {
    card: (
      <Stack tokens={{ childrenGap: 8 }}>
        {renderShimmer([
          { type: ShimmerElementType.circle, height: 48 },
          { type: ShimmerElementType.gap, height: 48, width: 16 },
          { type: ShimmerElementType.line, height: 16, width: 200 },
        ])}
        {renderShimmer([{ type: ShimmerElementType.line, height: 32 }])}
        {renderShimmer([{ type: ShimmerElementType.line, height: 16, width: '75%' }])}
      </Stack>
    ),

    list: (
      <Stack tokens={{ childrenGap: 12 }}>
        {Array.from({ length: count }).map((_, i) => (
          <Stack key={i} horizontal tokens={{ childrenGap: 12 }} verticalAlign="center">
            <Shimmer
              shimmerElements={[{ type: ShimmerElementType.circle, height: 32 }]}
              width={32}
            />
            <Stack grow tokens={{ childrenGap: 4 }}>
              <Shimmer shimmerElements={[{ type: ShimmerElementType.line, height: 16 }]} />
              <Shimmer shimmerElements={[{ type: ShimmerElementType.line, height: 12, width: '60%' }]} />
            </Stack>
          </Stack>
        ))}
      </Stack>
    ),

    detail: (
      <Stack tokens={{ childrenGap: 16 }}>
        {renderShimmer([
          { type: ShimmerElementType.line, height: 32, width: 300 },
        ])}
        <Stack horizontal tokens={{ childrenGap: 16 }}>
          {renderShimmer([
            { type: ShimmerElementType.circle, height: 80 },
            { type: ShimmerElementType.gap, height: 80, width: 16 },
            { type: ShimmerElementType.line, height: 80, width: 200 },
          ])}
        </Stack>
        {Array.from({ length: 4 }).map((_, i) => (
          <Shimmer key={i} shimmerElements={[{ type: ShimmerElementType.line, height: 16 }]} />
        ))}
      </Stack>
    ),

    form: (
      <Stack tokens={{ childrenGap: 16 }}>
        {Array.from({ length: count }).map((_, i) => (
          <Stack key={i} tokens={{ childrenGap: 4 }}>
            <Shimmer
              shimmerElements={[{ type: ShimmerElementType.line, height: 14, width: 100 }]}
              width={100}
            />
            <Shimmer shimmerElements={[{ type: ShimmerElementType.line, height: 32 }]} />
          </Stack>
        ))}
      </Stack>
    ),

    table: (
      <Stack tokens={{ childrenGap: 8 }}>
        {/* Header */}
        <Shimmer
          shimmerElements={[
            { type: ShimmerElementType.line, height: 32 },
          ]}
        />
        {/* Rows */}
        {Array.from({ length: count }).map((_, i) => (
          <Stack key={i} horizontal tokens={{ childrenGap: 16 }}>
            {Array.from({ length: 4 }).map((__, j) => (
              <Stack.Item key={j} grow>
                <Shimmer shimmerElements={[{ type: ShimmerElementType.line, height: 24 }]} />
              </Stack.Item>
            ))}
          </Stack>
        ))}
      </Stack>
    ),

    text: (
      <Stack tokens={{ childrenGap: 8 }}>
        {Array.from({ length: count }).map((_, i) => (
          <Shimmer
            key={i}
            shimmerElements={[
              {
                type: ShimmerElementType.line,
                height: 16,
                width: i === count - 1 ? '60%' : '100%',
              },
            ]}
          />
        ))}
      </Stack>
    ),

    custom: customElements ? renderShimmer(customElements) : null,
  };

  return (
    <Stack
      tokens={{ padding: 16 }}
      styles={{ root: { opacity: 0.8 } }}
    >
      {message && (
        <Text styles={{ root: { color: '#605e5c', marginBottom: 16 } }}>
          {message}
        </Text>
      )}
      {patterns[type]}
    </Stack>
  );
};

// Convenience components for common patterns
export const LoadingPatterns = {
  /** Card loading pattern */
  Card: (props: Omit<LoadingStateProps, 'type'>) => <LoadingState type="card" {...props} />,
  
  /** List loading pattern */
  List: (props: Omit<LoadingStateProps, 'type'>) => <LoadingState type="list" {...props} />,
  
  /** Detail view loading pattern */
  Detail: (props: Omit<LoadingStateProps, 'type'>) => <LoadingState type="detail" {...props} />,
  
  /** Form loading pattern */
  Form: (props: Omit<LoadingStateProps, 'type'>) => <LoadingState type="form" {...props} />,
  
  /** Table loading pattern */
  Table: (props: Omit<LoadingStateProps, 'type'>) => <LoadingState type="table" {...props} />,
  
  /** Text content loading pattern */
  Text: (props: Omit<LoadingStateProps, 'type'>) => <LoadingState type="text" {...props} />,
  
  /** Spinner only */
  Spinner: ({ message, height }: { message?: string; height?: number }) => (
    <LoadingState type="custom" useSpinner message={message} height={height} />
  ),
};

export default LoadingState;
