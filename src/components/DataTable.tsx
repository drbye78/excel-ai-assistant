// Data Table Component - Full-featured table with sorting, filtering, pagination
// For RecipeGallery, ConversationHistory, and other tabular data

import React, { useState, useMemo, useCallback } from 'react';
import {
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  SelectionMode,
  IColumn,
  Stack,
  Text,
  SearchBox,
  Dropdown,
  IDropdownOption,
  DefaultButton,
  IconButton,
  Label,
  mergeStyleSets,
} from '@fluentui/react';
import { EmptyState } from './EmptyState';
import { LoadingState } from './LoadingState';

export interface DataTableColumn<T = any> extends IColumn {
  /** Enable sorting for this column */
  isSortable?: boolean;
  /** Enable filtering for this column */
  isFilterable?: boolean;
  /** Custom filter options */
  filterOptions?: IDropdownOption[];
  /** Custom render function */
  onRender?: (item: T, index?: number, column?: IColumn) => React.ReactNode;
}

export interface DataTableProps<T = any> {
  /** Table data */
  items: T[];
  /** Column definitions */
  columns: DataTableColumn<T>[];
  /** Loading state */
  isLoading?: boolean;
  /** Selection mode */
  selectionMode?: SelectionMode;
  /** Callback when selection changes */
  onSelectionChange?: (selectedItems: T[]) => void;
  /** Enable global search */
  enableSearch?: boolean;
  /** Search placeholder */
  searchPlaceholder?: string;
  /** Enable column filtering */
  enableFiltering?: boolean;
  /** Enable pagination */
  enablePagination?: boolean;
  /** Items per page options */
  pageSizeOptions?: number[];
  /** Default page size */
  defaultPageSize?: number;
  /** Empty state message */
  emptyMessage?: string;
  /** Row actions */
  rowActions?: Array<{
    key: string;
    text: string;
    icon: string;
    onClick: (item: T) => void;
  }>;
  /** Compact mode */
  compact?: boolean;
}

interface SortState {
  key: string;
  direction: 'asc' | 'desc';
}

/**
 * Data Table Component
 * 
 * Full-featured table with sorting, filtering, pagination, and selection.
 * 
 * Usage:
 * ```tsx
 * <DataTable
 *   items={recipes}
 *   columns={[
 *     { key: 'name', name: 'Name', fieldName: 'name', isSortable: true },
 *     { key: 'date', name: 'Created', fieldName: 'createdAt', isSortable: true },
 *   ]}
 *   enableSearch
 *   enablePagination
 *   selectionMode={SelectionMode.multiple}
 *   onSelectionChange={(items) => setSelectedItems(items)}
 * />
 * ```
 */
export function DataTable<T extends Record<string, any>>({
  items,
  columns,
  isLoading,
  selectionMode = SelectionMode.none,
  onSelectionChange,
  enableSearch,
  searchPlaceholder = 'Search...',
  enableFiltering,
  enablePagination,
  pageSizeOptions = [10, 25, 50, 100],
  defaultPageSize = 25,
  emptyMessage = 'No data available',
  rowActions,
  compact,
}: DataTableProps<T>) {
  // State
  const [searchQuery, setSearchQuery] = useState('');
  const [sortState, setSortState] = useState<SortState | null>(null);
  const [filters, setFilters] = useState<Record<string, string>>({});
  const [currentPage, setCurrentPage] = useState(0);
  const [pageSize, setPageSize] = useState(defaultPageSize);
  const [selection] = useState(() => new Selection({
    onSelectionChanged: () => {
      onSelectionChange?.(selection.getSelection() as T[]);
    },
  }));

  // Reset page when filters change
  React.useEffect(() => {
    setCurrentPage(0);
  }, [searchQuery, filters, sortState]);

  // Filter and sort data
  const processedItems = useMemo(() => {
    let result = [...items];

    // Apply search
    if (enableSearch && searchQuery) {
      const query = searchQuery.toLowerCase();
      result = result.filter((item) =>
        columns.some((col) => {
          const value = col.fieldName ? item[col.fieldName] : null;
          if (value == null) return false;
          return String(value).toLowerCase().includes(query);
        })
      );
    }

    // Apply column filters
    if (enableFiltering) {
      Object.entries(filters).forEach(([key, value]) => {
        if (value) {
          result = result.filter((item) => {
            const col = columns.find((c) => c.key === key);
            if (!col || !col.fieldName) return true;
            return String(item[col.fieldName]) === value;
          });
        }
      });
    }

    // Apply sorting
    if (sortState) {
      const col = columns.find((c) => c.key === sortState.key);
      if (col && col.fieldName) {
        result.sort((a, b) => {
          const aVal = a[col.fieldName!];
          const bVal = b[col.fieldName!];
          
          if (aVal == null && bVal == null) return 0;
          if (aVal == null) return 1;
          if (bVal == null) return -1;
          
          const comparison = aVal < bVal ? -1 : aVal > bVal ? 1 : 0;
          return sortState.direction === 'asc' ? comparison : -comparison;
        });
      }
    }

    return result;
  }, [items, searchQuery, filters, sortState, columns, enableSearch, enableFiltering]);

  // Paginate data
  const paginatedItems = useMemo(() => {
    if (!enablePagination) return processedItems;
    const start = currentPage * pageSize;
    return processedItems.slice(start, start + pageSize);
  }, [processedItems, currentPage, pageSize, enablePagination]);

  const totalPages = Math.ceil(processedItems.length / pageSize);

  // Handle sort
  const handleColumnClick = useCallback((col: DataTableColumn<T>) => {
    if (!col.isSortable) return;
    
    setSortState((current) => {
      if (current?.key === col.key) {
        return current.direction === 'asc' 
          ? { key: col.key, direction: 'desc' }
          : null;
      }
      return { key: col.key, direction: 'asc' };
    });
  }, []);

  // Enhanced columns with sort indicators and actions
  const enhancedColumns = useMemo((): IColumn[] => {
    const baseColumns = columns.map((col): IColumn => ({
      ...col,
      isSorted: sortState?.key === col.key,
      isSortedDescending: sortState?.key === col.key && sortState.direction === 'desc',
      onColumnClick: () => handleColumnClick(col),
      styles: {
        root: { 
          cursor: col.isSortable ? 'pointer' : 'default',
          selectors: {
            ':hover': col.isSortable ? { backgroundColor: '#f3f2f1' } : {},
          },
        },
      },
    }));

    // Add actions column if rowActions provided
    if (rowActions && rowActions.length > 0) {
      baseColumns.push({
        key: 'actions',
        name: 'Actions',
        minWidth: 100,
        maxWidth: 100,
        onRender: (item: T) => (
          <Stack horizontal tokens={{ childrenGap: 4 }}>
            {rowActions.map((action) => (
              <IconButton
                key={action.key}
                iconProps={{ iconName: action.icon }}
                title={action.text}
                onClick={(e) => {
                  e?.stopPropagation();
                  action.onClick(item);
                }}
              />
            ))}
          </Stack>
        ),
      });
    }

    return baseColumns;
  }, [columns, sortState, handleColumnClick, rowActions]);

  const styles = mergeStyleSets({
    container: {
      display: 'flex',
      flexDirection: 'column',
      height: '100%',
    },
    toolbar: {
      padding: '12px 16px',
      borderBottom: '1px solid #e1dfdd',
      display: 'flex',
      justifyContent: 'space-between',
      alignItems: 'center',
      gap: '16px',
      flexWrap: 'wrap',
    },
    filters: {
      display: 'flex',
      gap: '12px',
      flexWrap: 'wrap',
    },
    tableContainer: {
      flex: 1,
      overflow: 'auto',
    },
    footer: {
      padding: '12px 16px',
      borderTop: '1px solid #e1dfdd',
      display: 'flex',
      justifyContent: 'space-between',
      alignItems: 'center',
    },
  });

  if (isLoading) {
    return <LoadingState type="table" count={5} />;
  }

  if (items.length === 0) {
    return (
      <EmptyState
        icon="Table"
        title={emptyMessage}
        description="Add some data to see it here."
      />
    );
  }

  return (
    <Stack className={styles.container}>
      {/* Toolbar */}
      {(enableSearch || enableFiltering) && (
        <Stack className={styles.toolbar}>
          {enableSearch && (
            <SearchBox
              placeholder={searchPlaceholder}
              value={searchQuery}
              onChange={(_, value) => setSearchQuery(value || '')}
              styles={{ root: { width: 300 } }}
            />
          )}
          
          {enableFiltering && (
            <Stack horizontal tokens={{ childrenGap: 8 }} className={styles.filters}>
              {columns
                .filter((col) => col.isFilterable && col.filterOptions)
                .map((col) => (
                  <Dropdown
                    key={col.key}
                    placeholder={`Filter ${col.name}`}
                    options={[{ key: '', text: 'All' }, ...(col.filterOptions || [])]}
                    selectedKey={filters[col.key] || ''}
                    onChange={(_, option) =>
                      setFilters((prev) => ({
                        ...prev,
                        [col.key]: option?.key as string,
                      }))
                    }
                    styles={{ root: { width: 150 } }}
                  />
                ))}
            </Stack>
          )}
        </Stack>
      )}

      {/* Table */}
      <div className={styles.tableContainer}>
        <DetailsList
          items={paginatedItems}
          columns={enhancedColumns}
          selection={selection}
          selectionMode={selectionMode}
          layoutMode={DetailsListLayoutMode.justified}
          isHeaderVisible
          compact={compact}
          styles={{
            root: { border: 'none' },
            headerWrapper: { 
              backgroundColor: '#faf9f8',
              borderBottom: '1px solid #e1dfdd',
            },
          }}
        />
      </div>

      {/* Footer with pagination */}
      {enablePagination && (
        <Stack className={styles.footer}>
          <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
            Showing {currentPage * pageSize + 1} -{' '}
            {Math.min((currentPage + 1) * pageSize, processedItems.length)} of{' '}
            {processedItems.length}
          </Text>
          
          <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="center">
            <Dropdown
              label="Items per page:"
              selectedKey={pageSize}
              options={pageSizeOptions.map((size) => ({
                key: size,
                text: String(size),
              }))}
              onChange={(_, option) => setPageSize(option?.key as number)}
              styles={{ root: { width: 80 }, label: { marginRight: 8 } }}
            />
            
            <Stack horizontal tokens={{ childrenGap: 4 }}>
              <IconButton
                iconProps={{ iconName: 'ChevronLeft' }}
                disabled={currentPage === 0}
                onClick={() => setCurrentPage((p) => p - 1)}
                ariaLabel="Previous page"
              />
              <Text styles={{ root: { padding: '8px 12px' } }}>
                {currentPage + 1} of {totalPages || 1}
              </Text>
              <IconButton
                iconProps={{ iconName: 'ChevronRight' }}
                disabled={currentPage >= totalPages - 1}
                onClick={() => setCurrentPage((p) => p + 1)}
                ariaLabel="Next page"
              />
            </Stack>
          </Stack>
        </Stack>
      )}
    </Stack>
  );
}

export default DataTable;
