# Implementation Plan - Fixing Outstanding Issues

## Status: In Progress

### ✅ Completed (8 issues)

1. **Production URL Configuration** - Fixed webpack.config.js to use environment variables
2. **Batch Operations Undo Payload** - Implemented proper snapshot capture and restoration
3. **M Code Generator Error Handling** - Replaced TODO comments with proper error throwing
4. **Chart Service Funnel Chart** - Fixed placeholder range to use actual config data
5. **Chart Service Error Bars** - Implemented error bar support with fallback handling
6. **Conditional Command Engine** - Implemented proper evaluation for value comparison, cell state, and error checks
7. **Hardcoded URLs in Config** - Fixed logger and Settings component URLs
8. **Azure OpenAI Placeholder** - Updated to require user configuration

### 🔄 In Progress (2 issues)

1. **Manifest.xml URLs** - Build script created, needs testing
2. **Custom Functions LAMBDA** - Improved evaluation, but full Excel engine integration needed

### 📋 Remaining Tasks

#### Critical Production Blockers (3)
- [ ] Test and verify manifest.xml build script works correctly
- [ ] Update documentation URLs (example.com references)
- [ ] Verify all environment variable usage

#### Code Cleanup (3)
- [ ] Replace console.* statements with logger service (50+ instances)
- [ ] Remove or properly type @ts-ignore comments (20+ instances)
- [ ] Remove or move proba directory

#### Documentation (1)
- [ ] Document API limitations properly in code comments

### Implementation Notes

#### Console.log Replacement Strategy
- Replace `console.log` → `logger.info()`
- Replace `console.warn` → `logger.warn()`
- Replace `console.error` → `logger.error()`
- Files affected: ~15 service files

#### @ts-ignore Comments
- Most are for Office.js type limitations
- Create proper type definitions or use type assertions
- Files affected: visualHighlighter, powerQueryService, powerPivotService, daxMeasureService

#### LAMBDA Evaluation
- Current implementation handles basic arithmetic
- Full Excel formula evaluation requires Excel API or formula parser library
- Document limitation clearly

### Next Steps

1. Complete console.log replacement (high priority)
2. Test manifest build script
3. Update documentation
4. Create type definitions for Office.js gaps
5. Final testing and validation
