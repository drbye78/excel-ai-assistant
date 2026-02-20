// Navigation Shell - Unified sidebar navigation for all features
// Replaces fragmented tab-based navigation

import React, { useState } from 'react';
import {
  Nav,
  INavLinkGroup,
  INavLink,
  Stack,
  IconButton,
  Text,
  TooltipHost,
  ITooltipHostStyles,
} from '@fluentui/react';

export interface NavItem {
  key: string;
  label: string;
  icon: string;
  component: React.ComponentType<any>;
  badge?: number;
  disabled?: boolean;
  description?: string;
}

export interface NavigationShellProps {
  /** Navigation items configuration */
  items: NavItem[];
  /** Currently selected key */
  selectedKey: string;
  /** Callback when selection changes */
  onSelect: (key: string) => void;
  /** Whether sidebar is collapsed */
  collapsed?: boolean;
  /** Toggle collapsed state */
  onToggleCollapse?: () => void;
  /** Header component or title */
  header?: React.ReactNode;
  /** Footer component */
  footer?: React.ReactNode;
}

/**
 * Navigation Shell Component
 * 
 * Provides unified sidebar navigation exposing all features.
 * Replaces the limited 2-tab navigation in App.tsx.
 * 
 * Usage:
 * ```tsx
 * const navItems: NavItem[] = [
 *   { key: 'chat', label: 'AI Assistant', icon: 'Chat', component: Chat },
 *   { key: 'recipes', label: 'Recipes', icon: 'ClipboardList', component: RecipeGallery },
 *   { key: 'analytics', label: 'Analytics', icon: 'Chart', component: AnalyticsDashboard },
 *   { key: 'batch', label: 'Batch Operations', icon: 'Processing', component: BatchOperationsPanel },
 *   { key: 'history', label: 'History', icon: 'History', component: ConversationHistory },
 *   { key: 'settings', label: 'Settings', icon: 'Settings', component: Settings },
 * ];
 * 
 * <NavigationShell
 *   items={navItems}
 *   selectedKey={currentView}
 *   onSelect={setCurrentView}
 * />
 * ```
 */
export const NavigationShell: React.FC<NavigationShellProps> = ({
  items,
  selectedKey,
  onSelect,
  collapsed = false,
  onToggleCollapse,
  header,
  footer,
}) => {
  const [isCollapsed, setIsCollapsed] = useState(collapsed);

  const handleToggle = () => {
    const newState = !isCollapsed;
    setIsCollapsed(newState);
    onToggleCollapse?.();
  };

  // Convert NavItems to INavLink format
  const navGroups: INavLinkGroup[] = [
    {
      links: items.map((item) => ({
        key: item.key,
        name: item.label,
        icon: item.icon,
        url: '',
        disabled: item.disabled,
        title: item.description || item.label,
        ...(item.badge && {
          onRenderIcon: () => (
            <Stack horizontal verticalAlign="center">
              <span>{item.label}</span>
              <span
                style={{
                  marginLeft: 8,
                  backgroundColor: '#0078d4',
                  color: 'white',
                  borderRadius: 10,
                  padding: '0 6px',
                  fontSize: 12,
                  fontWeight: 600,
                }}
              >
                {item.badge}
              </span>
            </Stack>
          ),
        }),
      })),
    },
  ];

  const handleLinkClick = (ev?: React.MouseEvent<HTMLElement>, item?: INavLink) => {
    if (item && !item.disabled) {
      onSelect(item.key);
    }
  };

  const SelectedComponent = items.find((item) => item.key === selectedKey)?.component;

  return (
    <Stack horizontal styles={{ root: { height: '100vh', overflow: 'hidden' } }}>
      {/* Sidebar */}
      <Stack
        styles={{
          root: {
            width: isCollapsed ? 48 : 220,
            minWidth: isCollapsed ? 48 : 220,
            backgroundColor: '#f3f2f1',
            borderRight: '1px solid #e1dfdd',
            transition: 'width 0.2s ease-in-out',
            display: 'flex',
            flexDirection: 'column',
          },
        }}
      >
        {/* Header */}
        <Stack
          horizontal
          horizontalAlign="space-between"
          verticalAlign="center"
          styles={{ root: { padding: '12px 16px', borderBottom: '1px solid #e1dfdd' } }}
        >
          {!isCollapsed && header}
          <TooltipHost
            content={isCollapsed ? 'Expand sidebar' : 'Collapse sidebar'}
            styles={{ root: { display: 'flex' } }}
          >
            <IconButton
              iconProps={{ iconName: isCollapsed ? 'DoubleChevronRight' : 'DoubleChevronLeft' }}
              onClick={handleToggle}
              styles={{ root: { marginLeft: isCollapsed ? 0 : 'auto' } }}
              ariaLabel={isCollapsed ? 'Expand navigation' : 'Collapse navigation'}
            />
          </TooltipHost>
        </Stack>

        {/* Navigation */}
        <Stack.Item grow styles={{ root: { overflowY: 'auto' } }}>
          <Nav
            groups={navGroups}
            selectedKey={selectedKey}
            onLinkClick={handleLinkClick}
            isCollapsed={isCollapsed}
            styles={{
              root: {
                paddingTop: 8,
              },
              navItems: {
                margin: 0,
              },
              navItem: {
                selectors: {
                  '&:hover': {
                    backgroundColor: '#e1dfdd',
                  },
                  '&.is-selected': {
                    backgroundColor: '#edebe9',
                    borderLeft: '3px solid #0078d4',
                  },
                },
              },
            }}
          />
        </Stack.Item>

        {/* Footer */}
        {footer && (
          <Stack styles={{ root: { padding: 12, borderTop: '1px solid #e1dfdd' } }}>
            {footer}
          </Stack>
        )}
      </Stack>

      {/* Main Content Area */}
      <Stack.Item grow styles={{ root: { overflow: 'auto', backgroundColor: '#faf9f8' } }}>
        <Stack styles={{ root: { minHeight: '100%' } }}>
          {SelectedComponent ? (
            <SelectedComponent />
          ) : (
            <Stack
              horizontalAlign="center"
              verticalAlign="center"
              styles={{ root: { height: '100%', padding: 40 } }}
            >
              <Text variant="large" styles={{ root: { color: '#605e5c' } }}>
                Select a feature from the sidebar
              </Text>
            </Stack>
          )}
        </Stack>
      </Stack.Item>
    </Stack>
  );
};

// Convenience hook for navigation state
export const useNavigation = (items: NavItem[], defaultKey?: string) => {
  const [selectedKey, setSelectedKey] = useState(defaultKey || items[0]?.key);
  const [collapsed, setCollapsed] = useState(false);

  const selectedItem = items.find((item) => item.key === selectedKey);

  return {
    selectedKey,
    setSelectedKey,
    selectedItem,
    collapsed,
    setCollapsed,
    SelectedComponent: selectedItem?.component,
  };
};

// Default navigation configuration for Excel AI Assistant
export const defaultNavItems: NavItem[] = [
  {
    key: 'chat',
    label: 'AI Assistant',
    icon: 'Chat',
    component: () => <div>Chat Component</div>, // Replace with actual import
    description: 'Chat with AI to perform Excel operations',
  },
  {
    key: 'recipes',
    label: 'Recipes',
    icon: 'ClipboardList',
    component: () => <div>Recipe Gallery</div>, // Replace with actual import
    description: 'Saved operation recipes',
  },
  {
    key: 'analytics',
    label: 'Analytics',
    icon: 'Chart',
    component: () => <div>Analytics Dashboard</div>, // Replace with actual import
    description: 'Usage analytics and insights',
  },
  {
    key: 'batch',
    label: 'Batch Operations',
    icon: 'Processing',
    component: () => <div>Batch Operations</div>, // Replace with actual import
    description: 'Run operations on multiple files',
  },
  {
    key: 'history',
    label: 'History',
    icon: 'History',
    component: () => <div>Conversation History</div>, // Replace with actual import
    description: 'Past conversations and operations',
  },
  {
    key: 'settings',
    label: 'Settings',
    icon: 'Settings',
    component: () => <div>Settings</div>, // Replace with actual import
    description: 'Configure the add-in',
  },
];

export default NavigationShell;
