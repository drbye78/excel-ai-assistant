// Help Panel Component - Contextual help system
// Provides inline documentation and guidance

import React, { useState } from 'react';
import { logger } from '../utils/logger';
import {
  Panel,
  PanelType,
  Stack,
  Text,
  Link,
  IconButton,
  Pivot,
  PivotItem,
  DefaultButton,
  IButtonStyles,
} from '@fluentui/react';

export interface HelpSection {
  title: string;
  content: React.ReactNode;
  icon?: string;
}

export interface HelpContent {
  title: string;
  description?: string;
  sections: HelpSection[];
  relatedLinks?: Array<{
    text: string;
    href: string;
    external?: boolean;
  }>;
  videoTutorials?: Array<{
    title: string;
    thumbnail: string;
    duration: string;
    url: string;
  }>;
}

export interface HelpPanelProps {
  /** Feature identifier for loading content */
  feature: string;
  /** Whether panel is open */
  isOpen: boolean;
  /** Close callback */
  onDismiss: () => void;
  /** Custom help content (optional) */
  content?: HelpContent;
}

/**
 * Built-in help content for common features
 */
const builtInHelpContent: Record<string, HelpContent> = {
  powerquery: {
    title: 'Power Query Builder',
    description: 'Create and manage data transformation queries using M language',
    sections: [
      {
        title: 'Getting Started',
        icon: 'Lightbulb',
        content: (
          <Stack tokens={{ childrenGap: 8 }}>
            <Text>Power Query lets you connect, transform, and load data from various sources.</Text>
            <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
              Start by creating a new query or editing an existing one from the list.
            </Text>
          </Stack>
        ),
      },
      {
        title: 'M Code Basics',
        icon: 'Code',
        content: (
          <Stack tokens={{ childrenGap: 8 }}>
            <Text>M is the formula language used in Power Query.</Text>
            <ul style={{ margin: 0, paddingLeft: 20 }}>
              <li><code>let</code> - Define variables</li>
              <li><code>in</code> - Return result</li>
              <li><code>Table.TransformColumns</code> - Modify columns</li>
              <li><code>Table.FilterRows</code> - Filter data</li>
            </ul>
          </Stack>
        ),
      },
      {
        title: 'Common Patterns',
        icon: 'Pattern',
        content: (
          <Stack tokens={{ childrenGap: 8 }}>
            <Text variant="small">Filter to recent data:</Text>
            <code style={{ backgroundColor: '#f3f2f1', padding: 8, display: 'block' }}>
              {'Table.SelectRows(Source, each [Date] >= Date.AddDays(DateTime.Date(DateTime.LocalNow()), -30))'}
            </code>
          </Stack>
        ),
      },
    ],
    relatedLinks: [
      { text: 'M Language Reference', href: 'https://docs.microsoft.com/powerquery-m', external: true },
      { text: 'Power Query Documentation', href: 'https://docs.microsoft.com/power-query', external: true },
    ],
  },
  dax: {
    title: 'DAX Measure Builder',
    description: 'Create calculated measures using Data Analysis Expressions',
    sections: [
      {
        title: 'Creating Measures',
        icon: 'Calculator',
        content: (
          <Stack tokens={{ childrenGap: 8 }}>
            <Text>DAX measures are calculations used in PivotTables and reports.</Text>
            <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
              Measures are evaluated at query time, making them dynamic and efficient.
            </Text>
          </Stack>
        ),
      },
      {
        title: 'Common Functions',
        icon: 'Function',
        content: (
          <Stack tokens={{ childrenGap: 8 }}>
            <ul style={{ margin: 0, paddingLeft: 20 }}>
              <li><code>SUM()</code> - Add values</li>
              <li><code>AVERAGE()</code> - Calculate mean</li>
              <li><code>CALCULATE()</code> - Modify filter context</li>
              <li><code>FILTER()</code> - Return filtered table</li>
              <li><code>RELATED()</code> - Get related value</li>
            </ul>
          </Stack>
        ),
      },
    ],
    relatedLinks: [
      { text: 'DAX Function Reference', href: 'https://docs.microsoft.com/dax', external: true },
      { text: 'DAX Patterns', href: 'https://daxpatterns.com', external: true },
    ],
  },
  recipes: {
    title: 'Recipe System',
    description: 'Save and reuse your common Excel operations',
    sections: [
      {
        title: 'What are Recipes?',
        icon: 'Recipe',
        content: (
          <Stack tokens={{ childrenGap: 8 }}>
            <Text>Recipes are saved sequences of operations that you can run with a single click.</Text>
            <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
              Perfect for repetitive tasks like monthly reporting or data cleaning.
            </Text>
          </Stack>
        ),
      },
      {
        title: 'Creating Recipes',
        icon: 'Add',
        content: (
          <Stack tokens={{ childrenGap: 8 }}>
            <ol style={{ margin: 0, paddingLeft: 20 }}>
              <li>Perform your desired operations in Excel</li>
              <li>Open the Recipe Builder</li>
              <li>Name and describe your recipe</li>
              <li>Save for future use</li>
            </ol>
          </Stack>
        ),
      },
    ],
  },
  chat: {
    title: 'AI Assistant',
    description: 'Use natural language to control Excel',
    sections: [
      {
        title: 'Natural Language Commands',
        icon: 'Chat',
        content: (
          <Stack tokens={{ childrenGap: 8 }}>
            <Text>Simply type what you want to do in plain English or Russian.</Text>
            <Text variant="small" styles={{ root: { fontWeight: 600 } }}>Examples:</Text>
            <ul style={{ margin: '4px 0', paddingLeft: 20, color: '#605e5c' }}>
              <li>"Create a pivot table from Sales data"</li>
              <li>"Format column A as currency"</li>
              <li>"Создай диаграмму продаж по регионам"</li>
            </ul>
          </Stack>
        ),
      },
      {
        title: 'Tips for Best Results',
        icon: 'Lightbulb',
        content: (
          <Stack tokens={{ childrenGap: 8 }}>
            <ul style={{ margin: 0, paddingLeft: 20 }}>
              <li>Be specific about cell ranges</li>
              <li>Mention sheet names when needed</li>
              <li>Use column headers by name</li>
              <li>Specify chart types explicitly</li>
            </ul>
          </Stack>
        ),
      },
    ],
  },
};

/**
 * HelpPanel Component
 * 
 * Provides contextual help with documentation, examples, and tutorials.
 * 
 * Usage:
 * ```tsx
 * const [showHelp, setShowHelp] = useState(false);
 * 
 * <IconButton
 *   iconProps={{ iconName: 'Help' }}
 *   onClick={() => setShowHelp(true)}
 * />
 * 
 * <HelpPanel
 *   feature="powerquery"
 *   isOpen={showHelp}
 *   onDismiss={() => setShowHelp(false)}
 * />
 * ```
 */
export const HelpPanel: React.FC<HelpPanelProps> = ({
  feature,
  isOpen,
  onDismiss,
  content: customContent,
}) => {
  const content = customContent || builtInHelpContent[feature];

  if (!content) {
    return null;
  }

  return (
    <Panel
      isOpen={isOpen}
      onDismiss={onDismiss}
      headerText={content.title}
      type={PanelType.medium}
      closeButtonAriaLabel="Close help panel"
    >
      <Stack tokens={{ childrenGap: 20 }}>
        {/* Description */}
        {content.description && (
          <Text styles={{ root: { color: '#605e5c' } }}>
            {content.description}
          </Text>
        )}

        {/* Sections */}
        <Pivot>
          <PivotItem headerText="Guide" itemIcon="BookAnswers">
            <Stack tokens={{ childrenGap: 16 }} style={{ marginTop: 16 }}>
              {content.sections.map((section, index) => (
                <Stack key={index} tokens={{ childrenGap: 8 }}>
                  <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                    {section.icon && (
                      <span style={{ color: '#0078d4' }}>
                        {/* Icon would be rendered here */}
                      </span>
                    )}
                    <Text variant="large" styles={{ root: { fontWeight: 600 } }}>
                      {section.title}
                    </Text>
                  </Stack>
                  {section.content}
                </Stack>
              ))}
            </Stack>
          </PivotItem>

          {content.videoTutorials && content.videoTutorials.length > 0 && (
            <PivotItem headerText="Videos" itemIcon="Video">
              <Stack tokens={{ childrenGap: 12 }} style={{ marginTop: 16 }}>
                {content.videoTutorials.map((video, index) => (
                  <Stack key={index} horizontal tokens={{ childrenGap: 12 }}>
                    <img
                      src={video.thumbnail}
                      alt={video.title}
                      width={120}
                      height={68}
                      style={{ borderRadius: 4, objectFit: 'cover' }}
                    />
                    <Stack>
                      <Link href={video.url} target="_blank">
                        {video.title}
                      </Link>
                      <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                        {video.duration}
                      </Text>
                    </Stack>
                  </Stack>
                ))}
              </Stack>
            </PivotItem>
          )}

          {content.relatedLinks && content.relatedLinks.length > 0 && (
            <PivotItem headerText="Resources" itemIcon="Link">
              <Stack tokens={{ childrenGap: 12 }} style={{ marginTop: 16 }}>
                {content.relatedLinks.map((link, index) => (
                  <Link
                    key={index}
                    href={link.href}
                    target={link.external ? '_blank' : undefined}
                  >
                    {link.text}
                    {link.external && ' ↗'}
                  </Link>
                ))}
              </Stack>
            </PivotItem>
          )}
        </Pivot>

        {/* Feedback */}
        <Stack
          horizontal
          horizontalAlign="space-between"
          verticalAlign="center"
          styles={{ root: { borderTop: '1px solid #e1dfdd', paddingTop: 16, marginTop: 16 } }}
        >
          <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
            Was this helpful?
          </Text>
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <IconButton
              iconProps={{ iconName: 'ThumbUp' }}
              title="Yes, this was helpful"
              onClick={() => logger.info('User feedback: helpful')}
            />
            <IconButton
              iconProps={{ iconName: 'ThumbDown' }}
              title="No, this wasn't helpful"
              onClick={() => logger.info('User feedback: not helpful')}
            />
          </Stack>
        </Stack>
      </Stack>
    </Panel>
  );
};

/**
 * Help Button Component
 * Convenience component for triggering help panel
 */
export interface HelpButtonProps {
  feature: string;
  label?: string;
}

export const HelpButton: React.FC<HelpButtonProps> = ({ feature, label }) => {
  const [isOpen, setIsOpen] = useState(false);

  return (
    <>
      {label ? (
        <DefaultButton
          iconProps={{ iconName: 'Help' }}
          text={label}
          onClick={() => setIsOpen(true)}
        />
      ) : (
        <IconButton
          iconProps={{ iconName: 'Help' }}
          title="Get help"
          ariaLabel="Get help"
          onClick={() => setIsOpen(true)}
        />
      )}
      <HelpPanel
        feature={feature}
        isOpen={isOpen}
        onDismiss={() => setIsOpen(false)}
      />
    </>
  );
};

export default HelpPanel;
