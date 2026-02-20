// DAX Generator - AI-Powered DAX Generation from Natural Language
// Phase 4 Implementation - Natural Language to DAX

import aiService from './aiService';
import { daxParser } from './daxParser';
import { logger } from '../utils/logger';

export interface DAXGenerationRequest {
  description: string;
  context?: {
    availableTables?: Array<{
      name: string;
      columns: string[];
    }>;
    existingMeasures?: string[];
    dateTable?: string;
    dateColumn?: string;
  };
  measureType?: 'simple' | 'time_intelligence' | 'comparison' | 'ratio' | 'cumulative';
  targetTable?: string;
}

export interface DAXGenerationResult {
  expression: string;
  measureName: string;
  explanation: string;
  confidence: 'high' | 'medium' | 'low';
  alternatives?: string[];
  warnings?: string[];
}

export class DAXGenerator {
  private static instance: DAXGenerator;

  private constructor() {}

  static getInstance(): DAXGenerator {
    if (!DAXGenerator.instance) {
      DAXGenerator.instance = new DAXGenerator();
    }
    return DAXGenerator.instance;
  }

  /**
   * Generate DAX from natural language description
   */
  async generateDAX(request: DAXGenerationRequest): Promise<DAXGenerationResult> {
    logger.info('Generating DAX from natural language', { description: request.description.substring(0, 100) });

    try {
      // Try AI-powered generation first
      const aiResult = await this.generateWithAI(request);
      if (aiResult && aiResult.confidence === 'high') {
        return aiResult;
      }
    } catch (error) {
      logger.warn('AI generation failed, falling back to templates', { error });
    }

    // Fallback to template-based generation
    return this.generateFromTemplate(request);
  }

  /**
   * Generate DAX using AI service
   */
  private async generateWithAI(request: DAXGenerationRequest): Promise<DAXGenerationResult | null> {
    const settings = aiService.getSettings();

    if (!settings.apiKey) {
      logger.debug('No API key configured, skipping AI generation');
      return null;
    }

    const systemPrompt = `You are an expert DAX (Data Analysis Expressions) formula generator for Power Pivot and Power BI.

Rules:
1. Generate ONLY valid DAX code without markdown formatting or explanations
2. Use proper DAX syntax with single quotes for table names: 'Table Name'[Column]
3. Use DIVIDE() instead of / operator for safe division
4. Use variables (VAR) for complex calculations to improve readability
5. Return JSON format with: expression, measureName, explanation, confidence

Common DAX patterns:
- Sum: SUM('Table'[Column])
- Average: AVERAGE('Table'[Column])
- Count rows: COUNTROWS('Table')
- Distinct count: DISTINCTCOUNT('Table'[Column])
- Calculate with filter: CALCULATE([Measure], 'Table'[Column] = "Value")
- Year-to-date: TOTALYTD([Measure], 'Date'[Date])
- Year-over-year: 
  VAR Current = [Measure]
  VAR Prior = CALCULATE([Measure], SAMEPERIODLASTYEAR('Date'[Date]))
  RETURN DIVIDE(Current - Prior, Prior, 0)
- Running total:
  CALCULATE(
    [Measure],
    FILTER(
      ALL('Date'[Date]),
      'Date'[Date] <= MAX('Date'[Date])
    )
  )
- Filter context: ALL(), ALLEXCEPT(), VALUES(), FILTERS()
- Time intelligence: DATESYTD, DATESQTD, DATESMTD, SAMEPERIODLASTYEAR, PARALLELPERIOD

Response format:
{
  "measureName": "Name of the measure",
  "expression": "DAX expression here",
  "explanation": "Brief explanation of what this measure calculates",
  "confidence": "high|medium|low"
}`;

    const contextPrompt = request.context ? `
Available tables and columns:
${request.context.availableTables?.map(t => `- '${t.name}': ${t.columns.join(', ')}`).join('\n')}

Existing measures: ${request.context.existingMeasures?.join(', ') || 'None'}
Date table: ${request.context.dateTable || 'Not specified'}
Date column: ${request.context.dateColumn || 'Not specified'}
` : '';

    try {
      const response = await fetch(`${settings.apiUrl}/chat/completions`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${settings.apiKey}`
        },
        body: JSON.stringify({
          model: settings.model,
          messages: [
            { role: 'system', content: systemPrompt },
            { role: 'user', content: `Generate a DAX measure for: ${request.description}\n${contextPrompt}` }
          ],
          temperature: 0.3,
          max_tokens: 1000
        })
      });

      if (!response.ok) {
        throw new Error(`API Error: ${response.status}`);
      }

      const data = await response.json();
      const content = data.choices[0]?.message?.content?.trim() || '';

      // Try to parse JSON response
      try {
        const jsonMatch = content.match(/\{[\s\S]*\}/);
        if (jsonMatch) {
          const parsed = JSON.parse(jsonMatch[0]);

          // Validate the generated DAX
          const validation = daxParser.validate(parsed.expression);
          const isValid = validation.filter(e => e.severity === 'error').length === 0;

          if (isValid) {
            return {
              expression: parsed.expression,
              measureName: parsed.measureName || this.suggestMeasureName(request.description),
              explanation: parsed.explanation,
              confidence: parsed.confidence || 'high'
            };
          }
        }
      } catch (parseError) {
        // If JSON parsing fails, try to extract DAX directly
        const daxMatch = content.match(/=\s*([\s\S]+?)(?:\n|$)/);
        if (daxMatch) {
          const expression = daxMatch[1].trim();
          const validation = daxParser.validate(expression);
          const isValid = validation.filter(e => e.severity === 'error').length === 0;

          if (isValid) {
            return {
              expression,
              measureName: this.suggestMeasureName(request.description),
              explanation: 'Generated from natural language description',
              confidence: 'medium'
            };
          }
        }
      }

      return null;
    } catch (error) {
      logger.error('AI DAX generation error', { error });
      return null;
    }
  }

  /**
   * Generate DAX from templates (fallback)
   */
  private generateFromTemplate(request: DAXGenerationRequest): DAXGenerationResult {
    const desc = request.description.toLowerCase();
    const { context } = request;

    // Default values
    const tableName = context?.availableTables?.[0]?.name || 'Table';
    const columnName = context?.availableTables?.[0]?.columns[0] || 'Column';
    const dateTable = context?.dateTable || 'Date';
    const dateColumn = context?.dateColumn || 'Date';

    let expression = '';
    let measureName = '';
    let explanation = '';
    let confidence: 'high' | 'medium' | 'low' = 'medium';
    let warnings: string[] = [];

    // Simple aggregations
    if (desc.includes('sum') || desc.includes('total')) {
      measureName = `Total ${columnName}`;
      expression = `SUM('${tableName}'[${columnName}])`;
      explanation = `Calculates the sum of ${columnName} from the ${tableName} table`;
    }

    else if (desc.includes('average') || desc.includes('mean') || desc.includes('avg')) {
      measureName = `Average ${columnName}`;
      expression = `AVERAGE('${tableName}'[${columnName}])`;
      explanation = `Calculates the average of ${columnName}`;
    }

    else if (desc.includes('count') && desc.includes('distinct') || desc.includes('unique')) {
      measureName = `Unique ${columnName}`;
      expression = `DISTINCTCOUNT('${tableName}'[${columnName}])`;
      explanation = `Counts the number of unique values in ${columnName}`;
    }

    else if (desc.includes('count')) {
      measureName = `Row Count`;
      expression = `COUNTROWS('${tableName}')`;
      explanation = `Counts the number of rows in the ${tableName} table`;
    }

    else if (desc.includes('min') || desc.includes('minimum')) {
      measureName = `Min ${columnName}`;
      expression = `MIN('${tableName}'[${columnName}])`;
      explanation = `Returns the minimum value of ${columnName}`;
    }

    else if (desc.includes('max') || desc.includes('maximum')) {
      measureName = `Max ${columnName}`;
      expression = `MAX('${tableName}'[${columnName}])`;
      explanation = `Returns the maximum value of ${columnName}`;
    }

    // Time intelligence
    else if (desc.includes('year to date') || desc.includes('ytd')) {
      measureName = `YTD ${columnName}`;
      expression = `TOTALYTD(SUM('${tableName}'[${columnName}]), '${dateTable}'[${dateColumn}])`;
      explanation = `Calculates the year-to-date total of ${columnName}`;
    }

    else if (desc.includes('quarter to date') || desc.includes('qtd')) {
      measureName = `QTD ${columnName}`;
      expression = `TOTALQTD(SUM('${tableName}'[${columnName}]), '${dateTable}'[${dateColumn}])`;
      explanation = `Calculates the quarter-to-date total of ${columnName}`;
    }

    else if (desc.includes('month to date') || desc.includes('mtd')) {
      measureName = `MTD ${columnName}`;
      expression = `TOTALMTD(SUM('${tableName}'[${columnName}]), '${dateTable}'[${dateColumn}])`;
      explanation = `Calculates the month-to-date total of ${columnName}`;
    }

    else if (desc.includes('previous year') || desc.includes('last year') || desc.includes('prior year')) {
      measureName = `${columnName} PY`;
      expression = `CALCULATE(SUM('${tableName}'[${columnName}]), SAMEPERIODLASTYEAR('${dateTable}'[${dateColumn}]))`;
      explanation = `Calculates the ${columnName} for the same period in the previous year`;
    }

    else if (desc.includes('year over year') || desc.includes('yoy') || desc.includes('growth')) {
      measureName = `YoY Growth % ${columnName}`;
      expression = `VAR CurrentYear = SUM('${tableName}'[${columnName}])
VAR PreviousYear = CALCULATE(SUM('${tableName}'[${columnName}]), SAMEPERIODLASTYEAR('${dateTable}'[${dateColumn}]))
RETURN
    IF(PreviousYear <> 0, DIVIDE(CurrentYear - PreviousYear, PreviousYear, 0), 0)`;
      explanation = `Calculates the year-over-year growth percentage for ${columnName}`;
      confidence = 'high';
    }

    else if (desc.includes('running total') || desc.includes('cumulative') || desc.includes('accumulated')) {
      measureName = `Running Total ${columnName}`;
      expression = `CALCULATE(
    SUM('${tableName}'[${columnName}]),
    FILTER(
        ALL('${dateTable}'[${dateColumn}]),
        '${dateTable}'[${dateColumn}] <= MAX('${dateTable}'[${dateColumn}])
    )
)`;
      explanation = `Calculates the running (cumulative) total of ${columnName}`;
    }

    // Ratios and percentages
    else if (desc.includes('percentage of total') || desc.includes('percent of total') || desc.includes('% of total')) {
      measureName = `${columnName} % of Total`;
      expression = `VAR CurrentValue = SUM('${tableName}'[${columnName}])
VAR TotalValue = CALCULATE(SUM('${tableName}'[${columnName}]), ALL('${tableName}'))
RETURN
    DIVIDE(CurrentValue, TotalValue, 0)`;
      explanation = `Calculates the percentage of the total for ${columnName}`;
    }

    else if (desc.includes('market share') || desc.includes('share')) {
      measureName = `${columnName} Market Share`;
      expression = `VAR CurrentValue = SUM('${tableName}'[${columnName}])
VAR TotalValue = CALCULATE(SUM('${tableName}'[${columnName}]), ALL('${tableName}'))
RETURN
    DIVIDE(CurrentValue, TotalValue, 0)`;
      explanation = `Calculates the market share based on ${columnName}`;
    }

    // Comparison
    else if (desc.includes('compare') || desc.includes('versus') || desc.includes('vs')) {
      measureName = `${columnName} Comparison`;
      expression = `VAR Current = SUM('${tableName}'[${columnName}])
VAR Comparison = CALCULATE(SUM('${tableName}'[${columnName}]), ALL('${dateTable}'))
RETURN
    DIVIDE(Current - Comparison, Comparison, 0)`;
      explanation = `Compares current value to a baseline`;
      warnings.push('You may need to adjust the comparison logic based on your specific requirements');
    }

    // Filtered measures
    else if (desc.includes('filter') || desc.includes('where')) {
      const filterCol = this.extractColumnName(desc) || columnName;
      measureName = `${columnName} Filtered`;
      expression = `CALCULATE(
    SUM('${tableName}'[${columnName}]),
    '${tableName}'[${filterCol}] = "Value"
)`;
      explanation = `Calculates the sum of ${columnName} with a filter applied`;
      warnings.push('Replace "Value" with your actual filter value');
    }

    // Default fallback
    else {
      measureName = `${columnName} Total`;
      expression = `SUM('${tableName}'[${columnName}])`;
      explanation = `Basic sum of ${columnName} (template-based fallback)`;
      confidence = 'low';
      warnings.push('This is a basic template. Consider refining your description or using AI generation with an API key.');
    }

    return {
      expression,
      measureName,
      explanation,
      confidence,
      warnings
    };
  }

  /**
   * Suggest a measure name based on description
   */
  private suggestMeasureName(description: string): string {
    // Remove common words and extract key terms
    const words = description
      .toLowerCase()
      .replace(/\b(create|measure|calculate|of|the|a|an|for|from|in|on|by|with)\b/g, '')
      .replace(/[^\w\s]/g, '')
      .trim()
      .split(/\s+/)
      .filter(w => w.length > 0);

    // Capitalize and join
    return words
      .map(w => w.charAt(0).toUpperCase() + w.slice(1))
      .join(' ')
      .substring(0, 50) || 'New Measure';
  }

  /**
   * Extract potential column name from description
   */
  private extractColumnName(description: string): string | null {
    const patterns = [
      /column\s+(\w+)/i,
      /field\s+(\w+)/i,
      /by\s+(\w+)/i,
      /where\s+(\w+)/i,
    ];

    for (const pattern of patterns) {
      const match = description.match(pattern);
      if (match) return match[1];
    }

    return null;
  }

  /**
   * Generate multiple alternatives for a measure
   */
  async generateAlternatives(request: DAXGenerationRequest): Promise<string[]> {
    const alternatives: string[] = [];
    const desc = request.description.toLowerCase();

    // If it's a growth calculation, provide both VAR and non-VAR versions
    if (desc.includes('growth') || desc.includes('yoy') || desc.includes('change')) {
      const tableName = request.context?.availableTables?.[0]?.name || 'Table';
      const columnName = request.context?.availableTables?.[0]?.columns[0] || 'Column';
      const dateTable = request.context?.dateTable || 'Date';
      const dateColumn = request.context?.dateColumn || 'Date';

      alternatives.push(
        `DIVIDE([Current Year] - [Previous Year], [Previous Year], 0)`,
        `(SUM('${tableName}'[${columnName}]) - CALCULATE(SUM('${tableName}'[${columnName}]), SAMEPERIODLASTYEAR('${dateTable}'[${dateColumn}]))) / CALCULATE(SUM('${tableName}'[${columnName}]), SAMEPERIODLASTYEAR('${dateTable}'[${dateColumn}]))`
      );
    }

    // If it's a time intelligence, provide alternatives
    if (desc.includes('ytd') || desc.includes('year to date')) {
      const tableName = request.context?.availableTables?.[0]?.name || 'Table';
      const columnName = request.context?.availableTables?.[0]?.columns[0] || 'Column';
      const dateTable = request.context?.dateTable || 'Date';
      const dateColumn = request.context?.dateColumn || 'Date';

      alternatives.push(
        `CALCULATE(SUM('${tableName}'[${columnName}]), DATESYTD('${dateTable}'[${dateColumn}]))`,
        `CALCULATE(SUM('${tableName}'[${columnName}]), FILTER(ALL('${dateTable}'), '${dateTable}'[${dateColumn}] <= MAX('${dateTable}'[${dateColumn}]) && YEAR('${dateTable}'[${dateColumn}]) = YEAR(MAX('${dateTable}'[${dateColumn}]))))`
      );
    }

    return alternatives;
  }

  /**
   * Explain a DAX measure in plain English
   */
  explainDAX(expression: string): string {
    const explanation = daxParser.explain(expression);
    return explanation.summary;
  }

  /**
   * Validate and suggest fixes for a DAX expression
   */
  validateAndSuggest(expression: string): {
    isValid: boolean;
    errors: string[];
    suggestions: string[];
  } {
    const errors: string[] = [];
    const suggestions: string[] = [];

    // Syntax validation
    const syntaxErrors = daxParser.validate(expression);
    syntaxErrors.forEach(err => {
      if (err.severity === 'error') {
        errors.push(`${err.message} at position ${err.position}`);
      }
    });

    const isValid = errors.length === 0;

    if (isValid) {
      const parsed = daxParser.parse(expression);
      const explanation = daxParser.explain(expression);

      // Add optimization suggestions
      suggestions.push(...explanation.optimizations.map(o => o.suggestion));
      suggestions.push(...explanation.performanceHints);

      // Check for common best practices
      if (expression.includes('/') && !expression.includes('DIVIDE')) {
        suggestions.push('Use DIVIDE() instead of / operator to handle division by zero gracefully');
      }

      if (!expression.includes('VAR') && expression.length > 200) {
        suggestions.push('Consider using VAR to define variables for better readability');
      }
    }

    return { isValid, errors, suggestions };
  }
}

export default DAXGenerator.getInstance();
