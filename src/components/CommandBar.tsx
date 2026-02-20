// Command Bar Component - Contextual actions following Office design
// Provides consistent command patterns across the app

import React from 'react';
import {
  CommandBar as FluentCommandBar,
  ICommandBarItemProps,
  ICommandBarStyles,
} from '@fluentui/react';

export interface CommandBarProps {
  /** Primary action items */
  items: ICommandBarItemProps[];
  /** Far-right items (search, view toggle, etc.) */
  farItems?: ICommandBarItemProps[];
  /** Overflow items (shown when space is limited) */
  overflowItems?: ICommandBarItemProps[];
  /** Whether to show search box */
  showSearch?: boolean;
  /** Search placeholder text */
  searchPlaceholder?: string;
  /** Search callback */
  onSearch?: (query: string) => void;
  /** Custom styles */
  styles?: Partial<ICommandBarStyles>;
}

/**
 * Command Bar Component
 * 
 * Provides consistent contextual actions following Microsoft Office design patterns.
 * Replaces scattered buttons with organized command patterns.
 * 
 * Usage:
 * ```tsx
 * <CommandBar
 *   items={[
 *     { key: 'new', text: 'New', iconProps: { iconName: 'Add' }, onClick: () => {} },
 *     { key: 'edit', text: 'Edit', iconProps: { iconName: 'Edit' }, onClick: () => {} },
 *     { 
 *       key: 'delete', 
 *       text: 'Delete', 
 *       iconProps: { iconName: 'Delete' },
 *       onClick: () => {},
 *       disabled: !selection
 *     },
 *   ]}
 *   farItems={[
 *     { key: 'search', onRender: () => <SearchBox /> },
 *     { key: 'refresh', iconProps: { iconName: 'Refresh' }, onClick: () => {} },
 *   ]}
 * />
 * ```
 */
export const CommandBar: React.FC<CommandBarProps> = ({
  items,
  farItems,
  overflowItems,
  showSearch,
  searchPlaceholder = 'Search...',
  onSearch,
  styles,
}) => {
  const commandBarStyles: Partial<ICommandBarStyles> = {
    root: {
      padding: '0 16px',
      borderBottom: '1px solid #e1dfdd',
      backgroundColor: '#ffffff',
      ...styles?.root,
    },
  };

  return (
    <FluentCommandBar
      items={items}
      farItems={farItems}
      overflowItems={overflowItems}
      styles={commandBarStyles}
      ariaLabel="Command bar"
      primaryGroupAriaLabel="Primary commands"
      farItemsGroupAriaLabel="Far commands"
      overflowButtonProps={{
        ariaLabel: 'More commands',
      }}
    />
  );
};

// Common command patterns
export const CommandPatterns = {
  /** Standard CRUD operations */
  crud: (options: {
    onNew: () => void;
    onEdit?: () => void;
    onDelete?: () => void;
    onDuplicate?: () => void;
    hasSelection?: boolean;
  }): ICommandBarItemProps[] => [
    {
      key: 'new',
      text: 'New',
      iconProps: { iconName: 'Add' },
      onClick: options.onNew,
    },
    ...(options.onEdit
      ? [
          {
            key: 'edit',
            text: 'Edit',
            iconProps: { iconName: 'Edit' },
            onClick: options.onEdit,
            disabled: !options.hasSelection,
          } as ICommandBarItemProps,
        ]
      : []),
    ...(options.onDuplicate
      ? [
          {
            key: 'duplicate',
            text: 'Duplicate',
            iconProps: { iconName: 'Copy' },
            onClick: options.onDuplicate,
            disabled: !options.hasSelection,
          } as ICommandBarItemProps,
        ]
      : []),
    ...(options.onDelete
      ? [
          {
            key: 'delete',
            text: 'Delete',
            iconProps: { iconName: 'Delete' },
            onClick: options.onDelete,
            disabled: !options.hasSelection,
            buttonStyles: {
              root: { color: '#d13438' },
              rootDisabled: { color: '#a19f9d' },
            },
          } as ICommandBarItemProps,
        ]
      : []),
  ],

  /** Import/Export operations */
  importExport: (options: {
    onImport: () => void;
    onExport: () => void;
    onExportAll?: () => void;
  }): ICommandBarItemProps[] => [
    {
      key: 'import',
      text: 'Import',
      iconProps: { iconName: 'Download' },
      onClick: options.onImport,
    },
    {
      key: 'export',
      text: 'Export',
      iconProps: { iconName: 'Upload' },
      subMenuProps: {
        items: [
          {
            key: 'exportSelected',
            text: 'Export Selected',
            onClick: options.onExport,
          },
          ...(options.onExportAll
            ? [
                {
                  key: 'exportAll',
                  text: 'Export All',
                  onClick: options.onExportAll,
                } as ICommandBarItemProps,
              ]
            : []),
        ],
      },
    },
  ],

  /** View options (grid/list, refresh) */
  viewOptions: (options: {
    viewMode: 'grid' | 'list';
    onToggleView: () => void;
    onRefresh: () => void;
    onFilter?: () => void;
    onSort?: () => void;
  }): ICommandBarItemProps[] => [
    ...(options.onFilter
      ? [
          {
            key: 'filter',
            text: 'Filter',
            iconProps: { iconName: 'Filter' },
            onClick: options.onFilter,
          } as ICommandBarItemProps,
        ]
      : []),
    ...(options.onSort
      ? [
          {
            key: 'sort',
            text: 'Sort',
            iconProps: { iconName: 'Sort' },
            onClick: options.onSort,
          } as ICommandBarItemProps,
        ]
      : []),
    {
      key: 'refresh',
      text: 'Refresh',
      iconProps: { iconName: 'Refresh' },
      onClick: options.onRefresh,
    },
    {
      key: 'view',
      text: options.viewMode === 'grid' ? 'List View' : 'Grid View',
      iconProps: { iconName: options.viewMode === 'grid' ? 'List' : 'GridView' },
      onClick: options.onToggleView,
    },
  ],
};

export default CommandBar;
