# Suggested Intent Actions and Targets for Excel AI Assistant

## Current State Analysis

### Current Intents (18)
- create, modify, delete, explain, format, analyze, filter, sort, calculate, refresh, export, import, compare, merge, validate, suggest, automate, protect

### Current Targets (19)
- pivot, chart, table, query, measure, range, worksheet, workbook, shape, image, formula, data, namedRange, comment, sparkline, slicer, validation, hyperlink, macro

## 🎯 Suggested New Intents (High Priority)

### 1. **'duplicate' / 'copy'**
**Purpose:** Copy existing elements (distinct from 'create' which generates new)
**Use Cases:**
- "Duplicate this worksheet"
- "Copy this chart to Sheet2"
- "Make a copy of this pivot table"
- "Clone this table structure"

**Keywords:** duplicate, copy, clone, replicate, make a copy

---

### 2. **'move' / 'reorder'**
**Purpose:** Move or reorder elements within workbook
**Use Cases:**
- "Move column B to position D"
- "Reorder worksheets alphabetically"
- "Move this chart to Sheet2"
- "Shift rows 5-10 down by 2 rows"

**Keywords:** move, reorder, shift, relocate, rearrange, reposition

---

### 3. **'hide' / 'show'**
**Purpose:** Hide/show rows, columns, worksheets, or elements
**Use Cases:**
- "Hide columns C through E"
- "Show all hidden rows"
- "Hide this worksheet"
- "Unhide Sheet2"

**Keywords:** hide, show, unhide, display, conceal, reveal

---

### 4. **'freeze'**
**Purpose:** Freeze panes for navigation
**Use Cases:**
- "Freeze the top row"
- "Freeze first column"
- "Freeze panes at cell B2"
- "Unfreeze all panes"

**Keywords:** freeze, unfreeze, lock panes, stick

---

### 5. **'find' / 'search'**
**Purpose:** Find and search operations
**Use Cases:**
- "Find all cells containing 'Sales'"
- "Search for formulas in this range"
- "Find cells with errors"
- "Locate all merged cells"

**Keywords:** find, search, locate, seek, discover

---

### 6. **'replace'**
**Purpose:** Find and replace operations
**Use Cases:**
- "Replace 'Q1' with 'Quarter 1' in column A"
- "Find and replace formulas with values"
- "Replace all #N/A errors with 0"

**Keywords:** replace, substitute, swap, change all

---

### 7. **'group' / 'ungroup'**
**Purpose:** Group/ungroup rows, columns, or data
**Use Cases:**
- "Group rows 5-10"
- "Ungroup all grouped columns"
- "Create outline groups for this data"
- "Collapse all groups"

**Keywords:** group, ungroup, outline, collapse, expand

---

### 8. **'convert'**
**Purpose:** Convert between formats/types
**Use Cases:**
- "Convert this table to a range"
- "Convert formulas to values"
- "Convert text to numbers"
- "Convert dates to text format"

**Keywords:** convert, transform, change type, cast

---

### 9. **'link'**
**Purpose:** Create links between sheets/workbooks
**Use Cases:**
- "Link cell A1 to Sheet2!B5"
- "Create external link to Budget.xlsx"
- "Update all links in this workbook"
- "Break links to external workbooks"

**Keywords:** link, connect, reference, break link, update link

---

### 10. **'optimize'**
**Purpose:** Performance and optimization operations
**Use Cases:**
- "Optimize this workbook for performance"
- "Remove unused styles"
- "Compress images in this sheet"
- "Clean up formatting"

**Keywords:** optimize, improve performance, clean up, compress, streamline

---

## 🎯 Suggested New Intents (Medium Priority)

### 11. **'backup' / 'save'**
**Purpose:** Save and backup operations
**Use Cases:**
- "Save a backup copy"
- "Save as PDF"
- "Create a snapshot"
- "Save with a new name"

**Keywords:** backup, save, snapshot, archive

---

### 12. **'undo' / 'redo'**
**Purpose:** Undo/redo operations
**Use Cases:**
- "Undo the last 3 actions"
- "Redo the formatting change"
- "Show undo history"

**Keywords:** undo, redo, revert, restore

---

### 13. **'highlight'**
**Purpose:** Highlight cells (distinct from format)
**Use Cases:**
- "Highlight all cells with errors"
- "Highlight duplicates in red"
- "Highlight cells above average"

**Keywords:** highlight, mark, emphasize, spotlight

---

### 14. **'share' / 'collaborate'**
**Purpose:** Sharing and collaboration
**Use Cases:**
- "Share this workbook with John"
- "Enable co-authoring"
- "Send this sheet via email"
- "Publish to SharePoint"

**Keywords:** share, collaborate, send, publish, distribute

---

### 15. **'schedule'**
**Purpose:** Schedule operations
**Use Cases:**
- "Schedule this macro to run daily"
- "Set up auto-refresh every hour"
- "Schedule data import at 9 AM"

**Keywords:** schedule, plan, automate timing, set up recurring

---

## 🎯 Suggested New Targets (High Priority)

### 1. **'row' / 'column'**
**Purpose:** Specific row/column operations
**Use Cases:**
- "Insert 3 rows above row 5"
- "Delete column B"
- "Resize column C to width 15"
- "Hide rows 10-20"

**Keywords:** row, column, col, rows, columns

---

### 2. **'cell'**
**Purpose:** Single cell operations
**Use Cases:**
- "Format cell A1 as currency"
- "Add comment to cell B5"
- "Link cell C10 to Sheet2!A1"
- "Clear cell D20"

**Keywords:** cell, cells

---

### 3. **'style'**
**Purpose:** Cell styles (distinct from format)
**Use Cases:**
- "Apply Heading 1 style to row 1"
- "Create a custom style called 'Highlight'"
- "Copy style from cell A1 to B1"
- "List all available styles"

**Keywords:** style, cell style, formatting style

---

### 4. **'connection'**
**Purpose:** Data connections
**Use Cases:**
- "Create connection to SQL Server"
- "Refresh all data connections"
- "List all external connections"
- "Remove connection to Database1"

**Keywords:** connection, data source, external data, link

---

### 5. **'relationship'**
**Purpose:** Data model relationships (Power Pivot)
**Use Cases:**
- "Create relationship between Sales and Customers"
- "Show all relationships in the data model"
- "Delete relationship between Tables A and B"

**Keywords:** relationship, data model, relation, link tables

---

### 6. **'group'**
**Purpose:** Grouped rows/columns
**Use Cases:**
- "Group rows 5-10"
- "Ungroup all groups"
- "Collapse group level 2"
- "Show outline symbols"

**Keywords:** group, outline, grouped rows, grouped columns

---

### 7. **'view'**
**Purpose:** Custom views
**Use Cases:**
- "Create a view called 'Print View'"
- "Switch to view 'Data Entry'"
- "Delete view 'Old View'"
- "List all views"

**Keywords:** view, custom view, saved view

---

### 8. **'scenario'**
**Purpose:** What-if scenarios
**Use Cases:**
- "Create scenario 'Best Case'"
- "Show scenario summary"
- "Switch to scenario 'Worst Case'"
- "Merge scenarios from Budget.xlsx"

**Keywords:** scenario, what-if, analysis scenario

---

### 9. **'goal'**
**Purpose:** Goal Seek operations
**Use Cases:**
- "Use goal seek to make B10 equal 1000 by changing B5"
- "Find value for cell C20 to achieve target in D20"

**Keywords:** goal seek, goal, target value

---

### 10. **'print'**
**Purpose:** Print settings and operations
**Use Cases:**
- "Set print area to A1:F50"
- "Add header 'Monthly Report'"
- "Set margins to narrow"
- "Print preview"

**Keywords:** print, printing, print area, print settings

---

## 🎯 Suggested New Targets (Medium Priority)

### 11. **'page'**
**Purpose:** Page setup
**Use Cases:**
- "Set page orientation to landscape"
- "Set paper size to A4"
- "Adjust page breaks"
- "Set print quality to high"

**Keywords:** page, page setup, page break, orientation

---

### 12. **'header' / 'footer'**
**Purpose:** Headers and footers
**Use Cases:**
- "Add header 'Company Name'"
- "Insert page number in footer"
- "Add date to header"
- "Remove footer"

**Keywords:** header, footer, page header, page footer

---

### 13. **'outline'**
**Purpose:** Data outline/grouping structure
**Use Cases:**
- "Create outline for this data"
- "Show outline symbols"
- "Auto outline this range"
- "Clear outline"

**Keywords:** outline, data outline, grouping structure

---

### 14. **'permission'**
**Purpose:** Permissions and access control
**Use Cases:**
- "Protect this sheet with password"
- "Allow users to edit range A1:B10"
- "Remove protection from Sheet2"
- "Set workbook permissions"

**Keywords:** permission, access, protect, security, lock

---

### 15. **'audit'**
**Purpose:** Audit trail and tracking
**Use Cases:**
- "Show change history"
- "Track changes in this workbook"
- "Highlight all changes"
- "Accept/reject changes"

**Keywords:** audit, track changes, history, revision

---

## 📊 Priority Matrix

### High Priority (Implement First)
**Intents:** duplicate, move, hide/show, freeze, find, replace, group/ungroup, convert, link, optimize
**Targets:** row/column, cell, style, connection, relationship, group, view, scenario, goal, print

**Rationale:** These cover the most common Excel operations that users frequently need.

### Medium Priority (Implement Second)
**Intents:** backup/save, undo/redo, highlight, share/collaborate, schedule
**Targets:** page, header/footer, outline, permission, audit

**Rationale:** These add valuable functionality but are less frequently used.

---

## 🔄 Combination Examples

### New Intent + Existing Target
- `duplicate` + `worksheet` → "Duplicate Sheet1"
- `move` + `column` → "Move column B to position D"
- `freeze` + `row` → "Freeze top row"
- `find` + `formula` → "Find all formulas"
- `convert` + `table` → "Convert table to range"

### Existing Intent + New Target
- `create` + `scenario` → "Create scenario 'Best Case'"
- `modify` + `style` → "Modify Heading 1 style"
- `delete` + `connection` → "Delete connection to Database1"
- `explain` + `relationship` → "Explain relationship between Sales and Customers"

### New Intent + New Target
- `duplicate` + `row` → "Duplicate row 5"
- `group` + `outline` → "Create outline groups"
- `optimize` + `workbook` → "Optimize workbook performance"
- `schedule` + `macro` → "Schedule macro to run daily"

---

## 📈 Impact Assessment

### Coverage Increase
- **Current:** 18 intents × 19 targets = 342 combinations
- **With High Priority Additions:** 28 intents × 29 targets = 812 combinations (+137%)
- **With All Additions:** 33 intents × 34 targets = 1,122 combinations (+228%)

### User Value
- **High Priority:** Addresses ~70% of common Excel operations
- **Medium Priority:** Adds advanced features for power users
- **Total:** Comprehensive coverage of Excel functionality

---

## 🛠️ Implementation Considerations

### 1. **Backward Compatibility**
- Ensure existing commands still work
- Add new intents/targets without breaking current functionality

### 2. **Pattern Matching**
- Add keywords to `intentPatterns` and `targetPatterns`
- Update Russian translations in `russianIntentPatterns` and `russianTargetPatterns`

### 3. **Service Integration**
- May need new services for some targets (e.g., `scenarioService`, `goalSeekService`)
- Extend existing services for new intents

### 4. **Testing**
- Add test cases for each new intent/target combination
- Verify natural language parsing accuracy

### 5. **Documentation**
- Update README with new command examples
- Add to help panel and onboarding

---

## 🎯 Recommended Implementation Order

### Phase 1: High-Value, Low-Complexity
1. `duplicate` intent (works with existing targets)
2. `move` intent (works with existing targets)
3. `hide/show` intent (works with existing targets)
4. `row`/`column` targets (works with existing intents)
5. `cell` target (works with existing intents)

### Phase 2: High-Value, Medium-Complexity
6. `freeze` intent
7. `find`/`replace` intents
8. `group`/`ungroup` intent
9. `style` target
10. `convert` intent

### Phase 3: Advanced Features
11. `connection` target
12. `relationship` target
13. `scenario` target
14. `goal` target
15. `optimize` intent

---

## 💡 Additional Suggestions

### Intent Modifiers
Consider adding intent modifiers for more precise control:
- `quick` - Quick operations (e.g., "quick format")
- `batch` - Batch operations (e.g., "batch delete")
- `smart` - AI-suggested operations (e.g., "smart format")

### Target Modifiers
Consider adding target modifiers:
- `selected` - Selected range/cells
- `all` - All instances (e.g., "all worksheets")
- `visible` - Only visible elements
- `filtered` - Only filtered data

---

## 📝 Summary

**Recommended Additions:**
- **10 High-Priority Intents:** duplicate, move, hide/show, freeze, find, replace, group/ungroup, convert, link, optimize
- **10 High-Priority Targets:** row/column, cell, style, connection, relationship, group, view, scenario, goal, print
- **5 Medium-Priority Intents:** backup/save, undo/redo, highlight, share/collaborate, schedule
- **5 Medium-Priority Targets:** page, header/footer, outline, permission, audit

**Total New Combinations:** ~780 additional command combinations

This would significantly expand the assistant's capabilities and cover most Excel operations users need.
