# Excel AI Assistant 🤖📊

A powerful AI-powered assistant for Excel 2016+ that enables natural language interaction with your workbooks. Features 342 command combinations, bilingual support, dark mode, and enterprise-grade capabilities.

## ✨ Features

### 🧠 Natural Language Interface
- **342 Command Combinations**: 18 intents × 19 targets for comprehensive Excel control
- **Bilingual Support**: Full English and Russian language support
- **Smart Suggestions**: AI-powered context-aware recommendations
- **Recipe System**: Save, share, and replay complex workflows

### 🎨 Modern UI/UX
- **13 Production Components**: ErrorBoundary, ToastNotification, EmptyState, LoadingState, NavigationShell, HelpPanel, CommandBar, ConfirmDialog, DataTable, OnboardingTour, and more
- **Dark Mode**: Full theme system with light/dark/excel themes
- **Responsive Design**: Mobile-friendly with breakpoint detection
- **Keyboard Shortcuts**: 8 default shortcuts for power users
- **Onboarding Tour**: Interactive first-time user experience

### 📊 Full Excel Element Support

| Element | Operations |
|---------|-----------|
| **Cells** | Read/Write values & formulas, formatting, validation |
| **Ranges** | Copy, clear, format, auto-fit, named ranges |
| **Tables** | Create, style, add rows, totals, delete |
| **Charts** | Create, customize, delete (all major types) |
| **Pivot Tables** | Create, refresh, configure fields, explain |
| **Data Validation** | Lists, ranges, dates, custom formulas |
| **Conditional Formatting** | Rules, color scales, data bars, icon sets |
| **Named Ranges** | Create, update, delete, scope management |
| **Hyperlinks** | Add, modify, remove, validate |
| **Comments** | Add, edit, delete, threaded discussions |
| **Worksheets** | Add, delete, rename, navigate, protect |

### 🔧 Advanced Excel Features

#### Power Query & M Code
- **M Code Generator**: Create Power Query transformations from natural language
- **M Code Explainer**: Understand existing M code in plain English
- **Power Query Builder**: Visual query construction interface

#### DAX & Power Pivot
- **DAX Measure Generator**: Create calculated measures from descriptions
- **DAX Explainer**: Explain complex DAX formulas
- **Power Pivot Integration**: Manage data models and relationships

#### Analytics & Intelligence
- **Data Analysis**: Statistical analysis, trend detection, outliers
- **Advanced Analytics**: Regression, forecasting, correlation analysis
- **Visual Highlighting**: Interactive data exploration with visual cues
- **Batch Operations**: Execute multiple operations efficiently

### 🛡️ Enterprise Features
- **Error Recovery**: Automatic retry with exponential backoff
- **Clarification Engine**: Handles ambiguous commands intelligently
- **User Learning**: Adapts to user preferences over time
- **Conditional Commands**: Smart conditional execution
- **Compliance Support**: Audit trails and data governance
- **Enterprise Auth**: SSO and multi-tenant support

### 🔗 API Compatibility
Works with any OpenAI-compatible API:
- OpenAI (GPT-4, GPT-3.5, GPT-4o)
- Azure OpenAI Service
- Local models (LM Studio, Ollama, text-generation-webui)
- Anthropic Claude (via compatible endpoints)
- Custom enterprise endpoints

---

## 🚀 Quick Start

### Prerequisites
- Node.js 18+ 
- Excel 2016+ (Windows, Mac, or Web)
- API key from OpenAI or compatible service

### Installation

```bash
# Clone the repository
git clone https://github.com/yourusername/excel-ai-assistant.git
cd excel-ai-assistant

# Install dependencies
npm install

# Copy environment file
copy .env.example .env  # Windows
cp .env.example .env    # Mac/Linux

# Edit .env with your API credentials
```

### Environment Configuration

Create a `.env` file:

```env
# OpenAI Configuration
OPENAI_API_URL=https://api.openai.com/v1
OPENAI_API_KEY=sk-your-api-key-here
OPENAI_MODEL=gpt-4

# Or Azure OpenAI
# OPENAI_API_URL=https://your-resource.openai.azure.com/openai/deployments/your-deployment
# OPENAI_API_KEY=your-azure-key
# OPENAI_MODEL=gpt-4

# Or Local Model (LM Studio/Ollama)
# OPENAI_API_URL=http://localhost:1234/v1
# OPENAI_API_KEY=not-needed
# OPENAI_MODEL=local-model
```

### Development

```bash
# Start development server (runs on https://localhost:3000)
npm run dev

# Or run webpack and server separately
npm run dev-server    # Terminal 1: Webpack dev server
npm run start:server  # Terminal 2: Express API server
```

### Build for Production

```bash
# Create optimized production build
npm run build

# Output in dist/ folder:
# - taskpane.html/js
# - commands.js
# - manifest.xml
# - assets/
```

---

## 📥 Sideload in Excel

### Windows Desktop

1. Open Excel
2. Go to **Insert** → **Add-ins** → **My Add-ins** → **Manage My Add-ins**
3. Select **Upload My Add-in**
4. Browse to `dist/manifest.xml`

### Mac

```bash
# Use the Office-Addin-Debugging tool
npx office-addin-debugging start manifest.xml
```

Or manually:
1. Open Excel
2. **Insert** → **Add-ins** → **My Add-ins** → **Manage My Add-ins**
3. Select **My Add-ins** → **Add from file**
4. Select `dist/manifest.xml`

### Excel Online

1. Go to https://excel.office.com
2. Click **Insert** → **Office Add-ins** → **Upload Add-in**
3. Upload `dist/manifest.xml`

### Shared Folder (Network Deployment)

1. Create a shared network folder
2. Copy `dist/` contents to the folder
3. Update `manifest.xml` SourceLocation URLs to point to the shared folder
4. In Excel: **File** → **Options** → **Trust Center** → **Trusted Add-in Catalogs**
5. Add the shared folder URL

---

## 📝 Usage Guide

### Basic Commands

The assistant understands natural language commands across 18 intent categories:

#### 📊 Data Operations
```
"Sort column A in descending order"
"Filter rows where status equals 'Active'"
"Find and replace 'Old' with 'New' in column B"
"Remove duplicates from the selected range"
```

#### 🧮 Formulas & Calculations
```
"Calculate the average of column B excluding zeros"
"Create a formula to find duplicates in column A"
"Sum all values where status is 'Completed'"
"Add a running total in column D"
```

#### 📋 Tables
```
"Convert A1:F50 to a table with blue style"
"Add a total row that sums column D"
"Add a new row with today's date and 100"
"Apply medium style 2 to the current table"
```

#### 📈 Charts
```
"Create a line chart of monthly sales in A1:B12"
"Make a pie chart showing category distribution"
"Add a scatter plot with trend line"
"Create a bar chart comparing Q1 and Q2"
```

#### 🔍 Pivot Tables
```
"Create a pivot table from A1:D100 with Region as rows and Sum of Sales"
"Add a slicer for the Category field"
"Change the pivot table to show averages instead of sums"
"Create a pivot chart from the current pivot table"
```

#### ✅ Data Validation
```
"Add dropdown to B2:B100 with Yes, No, Maybe"
"Require dates between 2020 and 2025 in column C"
"Only allow numbers greater than 0 in D2:D50"
"Create a custom validation that checks email format"
```

#### 🎨 Formatting
```
"Format these cells as currency with 2 decimals"
"Apply conditional formatting - highlight values above 100 in green"
"Make the header row bold with blue background"
"Auto-fit all columns in the current sheet"
```

#### 🏷️ Named Ranges
```
"Create a named range 'SalesData' for A1:D100"
"Update 'TaxRate' to reference cell E5"
"Delete the named range 'OldData'"
"List all named ranges in this workbook"
```

### Advanced Features

#### 🧪 Power Query / M Code
```
"Generate M code to combine Sheet1 and Sheet2"
"Explain this M code: let Source = Excel.CurrentWorkbook()..."
"Create a query that filters dates after 2024-01-01"
"Add a custom column that concatenates FirstName and LastName"
```

#### 📐 DAX / Power Pivot
```
"Create a DAX measure for year-over-year growth"
"Explain this DAX formula: CALCULATE(SUM(Sales[Amount]), PREVIOUSYEAR(Date[Date]))"
"Add a calculated column for profit margin"
"Create a measure that counts unique customers"
```

#### 🔗 Hyperlinks & Navigation
```
"Add a hyperlink to cell A1 linking to Sheet2!B5"
"Create a link to https://example.com in the selected cell"
"Remove all hyperlinks from column C"
"Add an email link to support@company.com"
```

#### 💬 Comments & Collaboration
```
"Add a comment to cell A1: 'Review needed by Friday'"
"Show all comments in this worksheet"
"Delete the comment on cell B5"
"Reply to the comment in cell C10"
```

#### 🔄 Recipes & Automation
```
"Save this sequence as a recipe called 'Monthly Report'"
"Run the 'Sales Dashboard' recipe"
"Share the 'Data Cleanup' recipe with my team"
"Show me all available recipes"
```

---

## 🏗️ Project Architecture

### Directory Structure

```
excel-ai-assistant/
├── src/
│   ├── components/           # 18+ React components
│   │   ├── Chat.tsx          # Main chat interface
│   │   ├── Settings.tsx      # Configuration panel
│   │   ├── CommandBar.tsx    # Quick action buttons
│   │   ├── NavigationShell.tsx # App navigation
│   │   ├── DataTable.tsx     # Data visualization
│   │   ├── RecipeBuilder.tsx # Recipe creation UI
│   │   ├── AnalyticsDashboard.tsx # Analytics view
│   │   ├── ErrorBoundary.tsx # Error handling
│   │   ├── ToastNotification.tsx # User feedback
│   │   ├── OnboardingTour.tsx # First-time guide
│   │   ├── EmptyState.tsx    # Empty state display
│   │   ├── LoadingState.tsx  # Loading indicators
│   │   ├── HelpPanel.tsx     # Contextual help
│   │   ├── ConfirmDialog.tsx # Confirmation dialogs
│   │   ├── DataAnalysis.tsx  # Analysis results
│   │   ├── DAXExplainer.tsx  # DAX explanation UI
│   │   ├── FormulaExplainer.tsx # Formula help
│   │   ├── PowerQueryBuilder.tsx # Query builder
│   │   ├── VisualHighlightPanel.tsx # Visual cues
│   │   ├── BatchOperationsPanel.tsx # Batch UI
│   │   ├── ConversationHistory.tsx # Chat history
│   │   └── index.ts          # Component exports
│   │
│   ├── services/             # 35+ business logic services
│   │   ├── aiService.ts      # AI API communication
│   │   ├── excelService.ts   # Excel API wrapper
│   │   ├── naturalLanguageCommandParser.ts # NL parsing
│   │   ├── actionHandler.ts  # Action execution
│   │   ├── smartSuggestionEngine.ts # Suggestions
│   │   ├── errorRecoveryEngine.ts # Error handling
│   │   ├── clarificationEngine.ts # Ambiguity resolution
│   │   ├── userLearningEngine.ts # Personalization
│   │   ├── conditionalCommandEngine.ts # Conditional logic
│   │   ├── recipeService.ts  # Recipe management
│   │   ├── chartService.ts   # Chart operations
│   │   ├── pivotTableService.ts # Pivot table ops
│   │   ├── validationService.ts # Data validation
│   │   ├── namedRangeService.ts # Named ranges
│   │   ├── commentService.ts # Comments
│   │   ├── hyperlinkService.ts # Hyperlinks
│   │   ├── macroService.ts   # Macro generation
│   │   ├── powerQueryService.ts # Power Query
│   │   ├── daxMeasureService.ts # DAX measures
│   │   ├── powerPivotService.ts # Power Pivot
│   │   ├── dataAnalysis.ts   # Statistical analysis
│   │   ├── advancedAnalytics.ts # Advanced analytics
│   │   ├── analyticsService.ts # Usage analytics
│   │   ├── batchOperations.ts # Batch processing
│   │   ├── conversationStorage.ts # Chat persistence
│   │   ├── conditionalFormatting.ts # Formatting
│   │   ├── sortingFilteringService.ts # Sort/filter
│   │   ├── visualHighlighter.ts # Visual highlights
│   │   ├── localAIService.ts # Local AI support
│   │   ├── enterpriseAuth.ts # Enterprise auth
│   │   ├── complianceService.ts # Compliance
│   │   ├── customFunctions.ts # Excel functions
│   │   ├── diagramService.ts # Diagram generation
│   │   └── ... (40+ total services)
│   │
│   ├── hooks/                # Custom React hooks
│   │   ├── useKeyboardShortcuts.ts # Keyboard shortcuts
│   │   └── useBreakpoint.ts  # Responsive breakpoints
│   │
│   ├── theme/                # Theme system
│   │   ├── tokens.ts         # Design tokens
│   │   ├── ThemeProvider.tsx # Theme context
│   │   └── index.ts          # Theme exports
│   │
│   ├── i18n/                 # Internationalization
│   │   ├── index.ts          # i18n setup
│   │   ├── translations.ts   # EN/RU translations
│   │   ├── localeService.ts  # Locale management
│   │   ├── excelFunctions.ts # Excel function names
│   │   └── types.ts          # i18n types
│   │
│   ├── types/                # TypeScript definitions
│   │   └── index.ts          # All type definitions
│   │
│   ├── utils/                # Utility functions
│   │   ├── cellReferenceParser.ts # Cell parsing
│   │   ├── formulaTokenizer.ts # Formula parsing
│   │   ├── errors.ts         # Error handling
│   │   ├── logger.ts         # Logging
│   │   └── notificationManager.ts # Notifications
│   │
│   ├── taskpane/             # Task pane entry
│   │   ├── index.tsx         # Entry point
│   │   ├── App.tsx           # Root component
│   │   └── taskpane.html     # HTML template
│   │
│   └── commands/             # Ribbon commands
│       ├── commands.ts       # Command handlers
│       └── commands.html     # Command HTML
│
├── assets/                   # Icons and images
├── dist/                     # Build output
├── manifest.xml              # Office add-in manifest
├── package.json              # Dependencies
├── tsconfig.json             # TypeScript config
├── webpack.config.js         # Build configuration
└── README.md                 # This file
```

### Technology Stack

| Layer | Technology |
|-------|------------|
| **Framework** | React 18 + TypeScript 5 |
| **UI Library** | Microsoft Fluent UI React 8 |
| **Office Integration** | Office.js (Excel JavaScript API) |
| **Build Tool** | Webpack 5 |
| **AI Integration** | OpenAI SDK + Azure OpenAI |
| **Testing** | Jest + React Testing Library |
| **Linting** | ESLint + TypeScript ESLint |

---

## 🎨 Customization

### Theme Configuration

The add-in supports three built-in themes:

```typescript
// src/theme/tokens.ts
const themes = {
  light: { /* Light theme colors */ },
  dark: { /* Dark theme colors */ },
  excel: { /* Excel-matched colors */ }
};
```

Users can switch themes in Settings → Appearance.

### Keyboard Shortcuts

Default shortcuts (customizable in Settings):

| Shortcut | Action |
|----------|--------|
| `Ctrl+Shift+A` | Open AI Assistant |
| `Ctrl+Enter` | Send message |
| `Ctrl+Shift+L` | Toggle theme |
| `Ctrl+Shift+H` | Show help |
| `Ctrl+Shift+R` | Run last recipe |
| `Ctrl+Shift+S` | Open settings |
| `Esc` | Close panel |

### Recipe Creation

Save complex workflows as reusable recipes:

1. Execute your workflow in chat
2. Click **"Save as Recipe"**
3. Name your recipe
4. Add description and tags
5. Share with team or keep private

---

## 🚢 Deployment Options

### 1. Microsoft AppSource (Public)

1. Build: `npm run build`
2. Test: `npm run validate` (validate manifest)
3. Package: Create ZIP with manifest + assets
4. Submit: [Partner Center](https://partner.microsoft.com)
5. Await validation (3-5 business days)

### 2. Private Organization (Centralized Deployment)

1. Host files on your web server (HTTPS required)
2. Update `manifest.xml` with production URLs
3. Deploy via Microsoft 365 Admin Center:
   - **Settings** → **Integrated apps** → **Upload custom apps**
   - Upload `manifest.xml`
   - Assign to users/groups

### 3. SharePoint Catalog

1. Build: `npm run build`
2. Go to SharePoint Admin Center → **Apps** → **App Catalog**
3. Upload `manifest.xml`
4. Deploy to organization

### 4. Network Share

1. Copy `dist/` to network share
2. Update manifest URLs
3. Add share as Trusted Catalog in Excel

---

## 🔧 Troubleshooting

### Common Issues

#### "Cannot find module" errors
```bash
npm install
# or
npm ci
```

#### Certificate errors during development
```bash
npx office-addin-dev-certs install
# Or manually trust the certificate in browser
```

#### API connection fails
- Verify API URL is correct (should end with `/v1` for OpenAI)
- Check API key is valid and has credits
- Ensure CORS is configured on your API
- Check firewall/proxy settings

#### Excel not loading add-in
- Verify Office.js is loaded (check browser console)
- Ensure `manifest.xml` is valid: `npm run validate`
- Check HTTPS certificate is trusted
- Try clearing Office cache:
  - Windows: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\
  - Mac: `~/Library/Containers/com.microsoft.Excel/Data/Library/Application Support/Microsoft/Office/16.0/Wef/`

#### Build fails
```bash
# Clear caches
npm run clean  # if available
rmdir /s /q node_modules  # Windows
rm -rf node_modules       # Mac/Linux
npm install
```

### Debug Mode

Enable debug logging:
```typescript
// In browser console (F12)
localStorage.setItem('EXCEL_AI_DEBUG', 'true');
localStorage.setItem('EXCEL_AI_LOG_LEVEL', 'verbose');
```

### Performance Tips

- For large workbooks (>10k rows), use batch operations
- Enable "Performance Mode" in settings for faster processing
- Use named ranges instead of cell references for better AI understanding

---

## 🛡️ Security

### Data Handling
- ✅ API keys stored in browser memory only (not persisted)
- ✅ All AI API communication encrypted (HTTPS)
- ✅ Excel data processed client-side only
- ✅ No data sent to third-party servers except configured AI API
- ✅ Optional: Local AI mode for air-gapped environments

### Enterprise Security
- SSO integration available
- Audit logging for compliance
- Data loss prevention (DLP) compatible
- GDPR/CCPA compliant data handling

---

## 📊 Stats & Telemetry

The add-in collects anonymous usage statistics (optional, can be disabled):

- Feature usage frequency
- Error rates and types
- Performance metrics
- No spreadsheet data or content is collected

View your analytics: **Settings** → **Analytics**

---

## 🤝 Contributing

### Development Workflow

1. Fork the repository
2. Create feature branch: `git checkout -b feature/amazing-feature`
3. Make changes following code style
4. Add tests for new functionality
5. Run linting: `npm run lint`
6. Run tests: `npm test`
7. Commit with conventional commits
8. Push and create Pull Request

### Code Style

- TypeScript strict mode enabled
- ESLint + Prettier for formatting
- Component naming: PascalCase
- Service naming: camelCase
- Hook naming: useCamelCase

### Testing

```bash
# Run unit tests
npm test

# Run with coverage
npm run test:coverage

# Run integration tests
npm run test:integration
```

---

## 📚 Additional Resources

### Documentation
- [Office Add-ins Documentation](https://docs.microsoft.com/en-us/office/dev/add-ins/)
- [Excel JavaScript API Reference](https://docs.microsoft.com/en-us/javascript/api/excel)
- [Fluent UI React](https://developer.microsoft.com/en-us/fluentui#/controls/web)

### Community
- [Stack Overflow: office-js](https://stackoverflow.com/questions/tagged/office-js)
- [Microsoft Q&A: Office Add-ins](https://docs.microsoft.com/en-us/answers/topics/office-addins-dev.html)

### Related Projects
- [Office-Addin-Scripts](https://github.com/OfficeDev/Office-Addin-Scripts)
- [Script Lab](https://github.com/OfficeDev/script-lab)

---

## 📄 License

MIT License - See [LICENSE](LICENSE) file for details

---

## 🙏 Acknowledgments

- Microsoft Office Platform Team
- Fluent UI React Contributors
- OpenAI and Open Source AI Community

---

**Happy Spreadsheeting! 🤖📊**

*Built with ❤️ for Excel power users*
