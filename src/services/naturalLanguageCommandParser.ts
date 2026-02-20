// Natural Language Command Parser - Unified Parser for All Excel Operations
// Phase 6 Implementation - Natural Language Interface

export type CommandIntent =
  | 'create'
  | 'modify'
  | 'delete'
  | 'explain'
  | 'format'
  | 'analyze'
  | 'filter'
  | 'sort'
  | 'calculate'
  | 'refresh'
  | 'export'
  | 'import'
  | 'compare'
  | 'merge'
  | 'validate'
  | 'suggest'
  | 'automate'
  | 'protect'
  // New high-priority intents
  | 'duplicate'
  | 'move'
  | 'hide'
  | 'show'
  | 'freeze'
  | 'find'
  | 'replace'
  | 'group'
  | 'ungroup'
  | 'convert'
  | 'link'
  | 'optimize'
  // Medium-priority intents
  | 'backup'
  | 'save'
  | 'undo'
  | 'redo'
  | 'highlight'
  | 'share'
  | 'schedule';

export type CommandTarget =
  | 'pivot'
  | 'chart'
  | 'table'
  | 'query'
  | 'measure'
  | 'range'
  | 'worksheet'
  | 'workbook'
  | 'shape'
  | 'image'
  | 'formula'
  | 'data'
  | 'namedRange'
  | 'comment'
  | 'sparkline'
  | 'slicer'
  | 'validation'
  | 'hyperlink'
  | 'macro'
  // New high-priority targets
  | 'row'
  | 'column'
  | 'cell'
  | 'style'
  | 'connection'
  | 'relationship'
  | 'group'
  | 'view'
  | 'scenario'
  | 'goal'
  | 'print'
  // Medium-priority targets
  | 'page'
  | 'header'
  | 'footer'
  | 'outline'
  | 'permission'
  | 'audit';

export interface ParsedCommand {
  originalText: string;
  intent: CommandIntent;
  target: CommandTarget;
  confidence: 'high' | 'medium' | 'low';
  parameters: Record<string, any>;
  constraints: string[];
  suggestions?: string[];
  alternatives?: string[];
}

export interface NLContext {
  selectedRange?: string;
  selectedWorksheet?: string;
  activeTable?: string;
  availableTables?: string[];
  availableColumns?: string[];
  dataType?: 'numeric' | 'text' | 'date' | 'mixed';
  rowCount?: number;
  columnCount?: number;
  selectedChart?: string;
  selectedPivot?: string;
  clipboardContent?: string;
  previousCommand?: ParsedCommand;
  commandHistory?: string[];
}

export type SupportedLocale = 'en' | 'ru';

export interface ConversationState {
  currentTopic?: string;
  pendingOperation?: Partial<ParsedCommand>;
  awaitingConfirmation?: boolean;
  suggestedCompletions?: string[];
  lastTarget?: CommandTarget;
  lastIntent?: CommandIntent;
  accumulatedParameters?: Record<string, any>;
}

export class NaturalLanguageCommandParser {
  private static instance: NaturalLanguageCommandParser;
  private conversationState: ConversationState = {};
  private currentLocale: SupportedLocale = 'en';

  // Russian intent keywords
  private russianIntentPatterns: Record<CommandIntent, string[]> = {
    'create': ['создать', 'сделать', 'добавить', 'новый', 'новая', 'вставить', 'сгенерировать', 'построить', 'создай', 'сделай', 'добавь', 'вставь', 'построй'],
    'modify': ['изменить', 'поменять', 'обновить', 'редактировать', 'настроить', 'установить', 'применить', 'форматировать', 'измени', 'поменяй', 'обнови', 'отредактируй', 'настрой', 'установи', 'примени'],
    'delete': ['удалить', 'убрать', 'очистить', 'стереть', 'исключить', 'удали', 'убери', 'очисти', 'сотри'],
    'explain': ['объяснить', 'описать', 'что такое', 'как работает', 'расскажи про', 'проанализировать', 'разобрать', 'объясни', 'опиши', 'расскажи', 'проанализируй', 'разбери'],
    'format': ['форматировать', 'стиль', 'цвет', 'шрифт', 'выравнивание', 'граница', 'тема', 'форматируй', 'раскрась', 'выдели'],
    'analyze': ['анализировать', 'вычислить', 'подсчитать', 'суммировать', 'статистика', 'проанализируй', 'вычисли', 'подсчитай', 'просуммируй'],
    'filter': ['фильтровать', 'показать только', 'скрыть', 'отобразить', 'где', 'выбрать', 'отфильтруй', 'покажи только', 'спрячь'],
    'sort': ['сортировать', 'порядок', 'упорядочить', 'ранжировать', 'отсортировать', 'отсортируй', 'упорядочь'],
    'calculate': ['вычислить', 'посчитать', 'сумма', 'среднее', 'количество', 'итог', 'вычисли', 'посчитай', 'сложи'],
    'refresh': ['обновить', 'перезагрузить', 'синхронизировать', 'обнови', 'перезагрузи'],
    'export': ['экспортировать', 'сохранить как', 'скачать', 'вывод', 'экспортируй', 'сохрани как', 'скачай'],
    'import': ['импортировать', 'загрузить', 'открыть', 'прочитать', 'импортируй', 'загрузи', 'открой'],
    'compare': ['сравнить', 'сопоставить', 'сравни', 'сопоставь', 'найти различия', 'покажи отличия'],
    'merge': ['объединить', 'слить', 'склеить', 'соединить', 'объедини', 'слей', 'склей'],
    'validate': ['проверить', 'валидировать', 'проверь', 'найти ошибки', 'контроль'],
    'suggest': ['предложить', 'рекомендовать', 'посоветуй', 'что делать', 'как лучше'],
    'automate': ['автоматизировать', 'записать макрос', 'создать макрос', 'автоматизируй'],
    'protect': ['защитить', 'блокировать', 'защити', 'запретить редактирование', 'установить пароль'],
    // New high-priority intents
    'duplicate': ['дублировать', 'копировать', 'клонировать', 'создать копию', 'скопировать', 'дублируй', 'копируй'],
    'move': ['переместить', 'передвинуть', 'сдвинуть', 'переставить', 'перемести', 'передвинь'],
    'hide': ['скрыть', 'спрятать', 'скрой', 'спрячь'],
    'show': ['показать', 'отобразить', 'раскрыть', 'покажи', 'отобрази'],
    'freeze': ['закрепить', 'зафиксировать', 'закрепи', 'зафиксируй'],
    'find': ['найти', 'искать', 'найди', 'ищи'],
    'replace': ['заменить', 'подменить', 'замени', 'подмени'],
    'group': ['группировать', 'сгруппировать', 'группируй', 'сгруппируй'],
    'ungroup': ['разгруппировать', 'разгруппируй'],
    'convert': ['преобразовать', 'конвертировать', 'преобразуй', 'конвертируй'],
    'link': ['связать', 'создать связь', 'ссылка', 'свяжи'],
    'optimize': ['оптимизировать', 'улучшить', 'оптимизируй', 'улучши'],
    // Medium-priority intents
    'backup': ['резервная копия', 'бэкап', 'сохранить копию', 'архив'],
    'save': ['сохранить', 'сохрани'],
    'undo': ['отменить', 'вернуть', 'отмени', 'верни'],
    'redo': ['повторить', 'повтори'],
    'highlight': ['выделить', 'подчеркнуть', 'выдели', 'подчеркни'],
    'share': ['поделиться', 'отправить', 'опубликовать', 'поделись', 'отправь'],
    'schedule': ['запланировать', 'расписание', 'запланируй']
  };

  // Russian target keywords
  private russianTargetPatterns: Record<CommandTarget, string[]> = {
    'pivot': ['сводная', 'сводная таблица', 'кросс-таблица'],
    'chart': ['диаграмма', 'график', 'чарт', 'круговая диаграмма', 'столбчатая диаграмма'],
    'table': ['таблица', 'умная таблица', 'список'],
    'query': ['запрос', 'павер квери', 'power query', 'm код', 'трансформация'],
    'measure': ['мера', 'вычисляемое поле', 'кпэ', 'показатель'],
    'range': ['диапазон', 'ячейки', 'выделение'],
    'worksheet': ['лист', 'рабочий лист', 'вкладка'],
    'workbook': ['книга', 'рабочая книга', 'файл'],
    'shape': ['фигура', 'рисунок', 'стрелка'],
    'image': ['изображение', 'картинка', 'фото'],
    'formula': ['формула', 'функция', 'вычисление'],
    'data': ['данные', 'значения', 'содержимое'],
    'namedRange': ['именованный диапазон', 'имя диапазона', 'названный диапазон'],
    'comment': ['комментарий', 'примечание', 'заметка'],
    'sparkline': ['спарклайн', 'мини-график', 'тренд'],
    'slicer': ['слайсер', 'фильтр', 'срез'],
    'validation': ['проверка данных', 'валидация', 'ограничение ввода'],
    'hyperlink': ['ссылка', 'гиперссылка', 'ссылка на'],
    'macro': ['макрос', 'автоматизация', 'запись макроса'],
    // New high-priority targets
    'row': ['строка', 'строки'],
    'column': ['столбец', 'колонка', 'столбцы', 'колонки'],
    'cell': ['ячейка', 'ячейки'],
    'style': ['стиль', 'стиль ячейки', 'формат стиля'],
    'connection': ['соединение', 'источник данных', 'внешние данные', 'подключение'],
    'relationship': ['связь', 'модель данных', 'отношение', 'связь таблиц'],
    'group': ['группа', 'структура', 'сгруппированные строки', 'сгруппированные столбцы'],
    'view': ['представление', 'пользовательское представление', 'сохраненное представление'],
    'scenario': ['сценарий', 'что если', 'анализ сценариев'],
    'goal': ['подбор параметра', 'цель', 'целевое значение'],
    'print': ['печать', 'область печати', 'настройки печати'],
    // Medium-priority targets
    'page': ['страница', 'настройка страницы', 'разрыв страницы', 'ориентация'],
    'header': ['заголовок', 'верхний колонтитул'],
    'footer': ['нижний колонтитул', 'подвал'],
    'outline': ['структура', 'структура данных', 'группировка'],
    'permission': ['разрешение', 'доступ', 'защита', 'безопасность', 'блокировка', 'пароль'],
    'audit': ['аудит', 'отслеживание изменений', 'история', 'ревизия']
  };

  // Bilingual synonyms for column resolution
  private columnSynonyms: Record<string, string[]> = {
    'revenue': ['sales', 'turnover', 'income', 'proceeds', 'продажи', 'выручка', 'доход'],
    'cost': ['expense', 'expenditure', 'spending', 'затраты', 'расходы'],
    'profit': ['margin', 'earnings', 'gain', 'прибыль', 'маржа'],
    'date': ['day', 'time', 'period', 'дата', 'день', 'время'],
    'customer': ['client', 'buyer', 'consumer', 'клиент', 'покупатель'],
    'product': ['item', 'goods', 'sku', 'товар', 'продукт'],
    'quantity': ['qty', 'amount', 'count', 'количество', 'число'],
    'price': ['cost', 'rate', 'value', 'цена', 'стоимость']
  };

  private constructor() {}

  static getInstance(): NaturalLanguageCommandParser {
    if (!NaturalLanguageCommandParser.instance) {
      NaturalLanguageCommandParser.instance = new NaturalLanguageCommandParser();
    }
    return NaturalLanguageCommandParser.instance;
  }

  setLocale(locale: SupportedLocale): void {
    this.currentLocale = locale;
  }

  getLocale(): SupportedLocale {
    return this.currentLocale;
  }

  /**
   * Parse a natural language command into a structured command
   */
  parseCommand(text: string, context?: NLContext, locale?: SupportedLocale): ParsedCommand {
    // Set locale if provided
    if (locale) {
      this.currentLocale = locale;
    }

    // Resolve referential expressions (this, that, it, same)
    const resolvedText = context ? this.resolveReferentialExpressions(text, context) : text;
    
    const normalizedText = this.normalizeText(resolvedText);

    // Check for follow-up commands
    if (this.isFollowUpCommand(normalizedText)) {
      return this.handleFollowUpCommand(normalizedText, context);
    }

    // Detect intent
    const intent = this.detectIntent(normalizedText);

    // Detect target
    const target = this.detectTarget(normalizedText);

    // Extract parameters based on intent and target
    const parameters = this.extractParameters(normalizedText, intent, target, context);

    // Apply context-aware defaults
    if (context) {
      this.applyContextDefaults(parameters, context, target, intent);
    }

    // Extract constraints
    const constraints = this.extractConstraints(normalizedText);

    // Calculate confidence
    const confidence = this.calculateConfidence(intent, target, parameters, normalizedText);

    // Generate suggestions based on context
    const suggestions = context ? this.generateSuggestions(intent, target, context) : undefined;

    // Update conversation state
    this.updateConversationState(intent, target, parameters);

    return {
      originalText: text,
      intent,
      target,
      confidence,
      parameters,
      constraints,
      suggestions
    };
  }

  /**
   * Resolve referential expressions like "this", "that", "it", "same"
   */
  private resolveReferentialExpressions(text: string, context: NLContext): string {
    let resolved = text;

    // "this" -> selected range/chart/pivot
    resolved = resolved.replace(/\bthis\s+(range|selection|table|chart|pivot)\b/gi, (match, p1) => {
      if (p1 === 'range' || p1 === 'selection') return context.selectedRange || 'selection';
      if (p1 === 'chart') return context.selectedChart || 'chart';
      if (p1 === 'pivot') return context.selectedPivot || 'pivot';
      if (p1 === 'table') return context.activeTable || 'table';
      return match;
    });

    // "that" -> previous command's range
    resolved = resolved.replace(/\bthat\s+(range|selection)\b/gi, () => 
      context.previousCommand?.parameters?.range || context.selectedRange || 'selection');

    // "it" -> previous command's target
    resolved = resolved.replace(/\bit\b/gi, () => 
      context.previousCommand?.target || 'range');

    // "the same" -> previous command's formatting/style
    resolved = resolved.replace(/\bthe\s+same\s+(formatting|style|formula)\b/gi, () => {
      const prevFormat = context.previousCommand?.parameters;
      if (prevFormat?.numberFormat) return prevFormat.numberFormat;
      if (prevFormat?.fillColor) return prevFormat.fillColor;
      return 'same formatting';
    });

    // Russian referential expressions
    resolved = resolved.replace(/\bэтот\s+(диапазон|график|таблица)\b/gi, (match, p1) => {
      if (p1 === 'диапазон') return context.selectedRange || 'диапазон';
      if (p1 === 'график') return context.selectedChart || 'график';
      if (p1 === 'таблица') return context.activeTable || 'таблица';
      return match;
    });

    resolved = resolved.replace(/\bтот\s+же\s+(формат|стиль)\b/gi, () => 
      context.previousCommand?.parameters?.numberFormat || 'тот же формат');

    return resolved;
  }

  /**
   * Check if command is a follow-up to previous command
   */
  private isFollowUpCommand(text: string): boolean {
    const followUpPatterns = [
      /^(and|also|then|next|now|additionally)/i,
      /^\s*(but|however|instead)/i,
      /^\s*(make it|change it|update it)/i,
      /^\s*(what about|how about)/i,
      /^(а|и|также|затем|теперь|давай)/i,
      /^(сделай|измени|обнови)/i,
      /^(а|и)\s+потом/i,
      /^(а|и)\s+затем/i
    ];
    return followUpPatterns.some(p => p.test(text));
  }

  /**
   * Handle follow-up command by inheriting from previous state
   */
  private handleFollowUpCommand(text: string, context?: NLContext): ParsedCommand {
    const baseCommand: Partial<ParsedCommand> = {
      intent: this.conversationState.lastIntent || 'modify',
      target: this.conversationState.lastTarget || 'range',
      parameters: { ...this.conversationState.accumulatedParameters }
    };

    // Parse the new command
    const additional = this.parseNewCommand(text, context);
    
    return {
      originalText: text,
      intent: additional.intent || baseCommand.intent!,
      target: additional.target || baseCommand.target!,
      confidence: additional.confidence,
      parameters: { ...baseCommand.parameters, ...additional.parameters },
      constraints: [...(baseCommand.constraints || []), ...(additional.constraints || [])],
      suggestions: additional.suggestions
    };
  }

  /**
   * Parse command without follow-up handling
   */
  private parseNewCommand(text: string, context?: NLContext): ParsedCommand {
    const normalizedText = this.normalizeText(text);
    const intent = this.detectIntent(normalizedText);
    const target = this.detectTarget(normalizedText);
    const parameters = this.extractParameters(normalizedText, intent, target, context);
    const constraints = this.extractConstraints(normalizedText);
    const confidence = this.calculateConfidence(intent, target, parameters, normalizedText);
    
    return {
      originalText: text,
      intent,
      target,
      confidence,
      parameters,
      constraints
    };
  }

  /**
   * Update conversation state after parsing
   */
  private updateConversationState(
    intent: CommandIntent,
    target: CommandTarget,
    parameters: Record<string, any>
  ): void {
    this.conversationState.lastIntent = intent;
    this.conversationState.lastTarget = target;
    this.conversationState.accumulatedParameters = {
      ...this.conversationState.accumulatedParameters,
      ...parameters
    };
  }

  /**
   * Apply context-aware defaults to parameters
   */
  private applyContextDefaults(
    parameters: Record<string, any>,
    context: NLContext,
    target: CommandTarget,
    intent: CommandIntent
  ): void {
    // Use selected range if no range specified
    if (!parameters.range && context.selectedRange) {
      parameters.range = context.selectedRange;
    }

    // Use active table if no table specified
    if (!parameters.tableName && context.activeTable && target === 'table') {
      parameters.tableName = context.activeTable;
    }

    // Resolve column names using smart resolution
    if (parameters.columns && context.availableColumns) {
      parameters.columns = parameters.columns.map((col: string) =>
        this.resolveColumnName(col, context.availableColumns!) || col
      );
    }

    // Use selected chart/pivot for relevant targets
    if (target === 'chart' && context.selectedChart && !parameters.chartName) {
      parameters.chartName = context.selectedChart;
    }
    if (target === 'pivot' && context.selectedPivot && !parameters.pivotName) {
      parameters.pivotName = context.selectedPivot;
    }
  }

  /**
   * Smart column name resolution with synonyms and fuzzy matching
   */
  resolveColumnName(userInput: string, availableColumns: string[], threshold: number = 0.6): string | null {
    const normalizedInput = userInput.toLowerCase().trim();

    // Direct match
    const directMatch = availableColumns.find(
      col => col.toLowerCase() === normalizedInput
    );
    if (directMatch) return directMatch;

    // Synonym match (bilingual)
    for (const [canonical, variants] of Object.entries(this.columnSynonyms)) {
      const allVariants = [canonical, ...variants];
      
      if (allVariants.includes(normalizedInput)) {
        const match = availableColumns.find(col => {
          const lowerCol = col.toLowerCase();
          return lowerCol.includes(canonical) ||
                 variants.some(v => lowerCol.includes(v));
        });
        if (match) return match;
      }
    }

    // Fuzzy match using Levenshtein distance
    let bestMatch: string | null = null;
    let bestScore = 0;

    for (const col of availableColumns) {
      const score = this.calculateSimilarity(normalizedInput, col.toLowerCase());
      if (score > bestScore && score >= threshold) {
        bestScore = score;
        bestMatch = col;
      }
    }

    return bestMatch;
  }

  /**
   * Calculate Levenshtein similarity between two strings
   */
  private calculateSimilarity(a: string, b: string): number {
    const matrix: number[][] = [];
    for (let i = 0; i <= b.length; i++) {
      matrix[i] = [i];
    }
    for (let j = 0; j <= a.length; j++) {
      matrix[0][j] = j;
    }
    for (let i = 1; i <= b.length; i++) {
      for (let j = 1; j <= a.length; j++) {
        const cost = b[i - 1] === a[j - 1] ? 0 : 1;
        matrix[i][j] = Math.min(
          matrix[i - 1][j] + 1,
          matrix[i][j - 1] + 1,
          matrix[i - 1][j - 1] + cost
        );
      }
    }
    const distance = matrix[b.length][a.length];
    return 1 - distance / Math.max(a.length, b.length);
  }

  /**
   * Normalize text for parsing
   */
  private normalizeText(text: string): string {
    return text
      .toLowerCase()
      .replace(/[.,!?;]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  /**
   * Detect the intent of the command
   */
  private detectIntent(text: string): CommandIntent {
    const intentPatterns: Record<CommandIntent, string[]> = {
      'create': ['create', 'make', 'add', 'new', 'insert', 'generate', 'build'],
      'modify': ['modify', 'change', 'update', 'edit', 'adjust', 'set', 'apply', 'format'],
      'delete': ['delete', 'remove', 'clear', 'erase', 'eliminate', 'get rid of'],
      'explain': ['explain', 'describe', 'what is', 'how does', 'tell me about', 'analyze'],
      'format': ['format', 'style', 'color', 'font', 'align', 'border', 'theme'],
      'analyze': ['analyze', 'calculate', 'compute', 'summarize', 'aggregate', 'statistic'],
      'filter': ['filter', 'show only', 'hide', 'display', 'where', 'select'],
      'sort': ['sort', 'order', 'arrange', 'organize', 'rank'],
      'calculate': ['calculate', 'compute', 'sum', 'average', 'count', 'total'],
      'refresh': ['refresh', 'update', 'reload', 'sync', 'refresh data'],
      'export': ['export', 'save as', 'download', 'output', 'write to'],
      'import': ['import', 'load', 'open', 'read', 'bring in'],
      'compare': ['compare', 'diff', 'difference', 'versus', 'vs', 'contrast'],
      'merge': ['merge', 'combine', 'join', 'unify', 'consolidate'],
      'validate': ['validate', 'check', 'verify', 'audit', 'find errors'],
      'suggest': ['suggest', 'recommend', 'advise', 'what should', 'how to improve'],
      'automate': ['automate', 'record macro', 'create macro', 'script'],
      'protect': ['protect', 'lock', 'secure', 'prevent edit', 'password'],
      // New high-priority intents
      'duplicate': ['duplicate', 'copy', 'clone', 'replicate', 'make a copy', 'make copy'],
      'move': ['move', 'reorder', 'shift', 'relocate', 'rearrange', 'reposition'],
      'hide': ['hide', 'conceal', 'make invisible'],
      'show': ['show', 'unhide', 'display', 'reveal', 'make visible'],
      'freeze': ['freeze', 'unfreeze', 'lock panes', 'stick', 'pin'],
      'find': ['find', 'search', 'locate', 'seek', 'discover'],
      'replace': ['replace', 'substitute', 'swap', 'change all', 'find and replace'],
      'group': ['group', 'outline', 'collapse', 'create group'],
      'ungroup': ['ungroup', 'remove group', 'expand', 'clear outline'],
      'convert': ['convert', 'transform', 'change type', 'cast', 'change format'],
      'link': ['link', 'connect', 'reference', 'create link', 'make link'],
      'optimize': ['optimize', 'improve performance', 'clean up', 'compress', 'streamline'],
      // Medium-priority intents
      'backup': ['backup', 'back up', 'save copy', 'archive'],
      'save': ['save', 'save as', 'store', 'keep'],
      'undo': ['undo', 'revert', 'restore', 'go back'],
      'redo': ['redo', 'repeat', 'reapply'],
      'highlight': ['highlight', 'mark', 'emphasize', 'spotlight'],
      'share': ['share', 'collaborate', 'send', 'publish', 'distribute'],
      'schedule': ['schedule', 'plan', 'automate timing', 'set up recurring', 'set schedule']
    };

    // Merge with Russian patterns if needed
    if (this.currentLocale === 'ru' || this.containsRussian(text)) {
      for (const intent of Object.keys(intentPatterns) as CommandIntent[]) {
        intentPatterns[intent] = [...intentPatterns[intent], ...this.russianIntentPatterns[intent]];
      }
    }

    for (const [intent, patterns] of Object.entries(intentPatterns)) {
      for (const pattern of patterns) {
        if (text.includes(pattern)) {
          return intent as CommandIntent;
        }
      }
    }

    // Default to 'suggest' instead of 'explain' - this is more neutral
    // and allows the system to ask for clarification rather than giving explanations
    return 'suggest';
  }

  /**
   * Check if text contains Russian characters
   */
  private containsRussian(text: string): boolean {
    return /[\u0400-\u04FF]/.test(text);
  }

  /**
   * Detect the target of the command
   */
  private detectTarget(text: string): CommandTarget {
    const targetPatterns: Record<CommandTarget, string[]> = {
      'pivot': ['pivot', 'pivot table', 'crosstab', 'pivot chart'],
      'chart': ['chart', 'graph', 'plot', 'visualization', 'pie chart', 'bar chart', 'line chart'],
      'table': ['table', 'data table', 'excel table', 'list'],
      'query': ['query', 'power query', 'm code', 'transformation', 'etl'],
      'measure': ['measure', 'dax', 'calculated field', 'kpi', 'metric'],
      'range': ['range', 'cells', 'selection', 'a1', 'b2'],
      'worksheet': ['sheet', 'worksheet', 'tab', 'workbook page'],
      'workbook': ['workbook', 'file', 'excel file', 'spreadsheet'],
      'shape': ['shape', 'drawing', 'diagram', 'arrow', 'box', 'circle'],
      'image': ['image', 'picture', 'photo', 'icon', 'logo'],
      'formula': ['formula', 'function', 'calculation', 'equation'],
      'data': ['data', 'values', 'content', 'information'],
      'namedRange': ['named range', 'name range', 'define name', 'named cell'],
      'comment': ['comment', 'note', 'annotation'],
      'sparkline': ['sparkline', 'mini chart', 'spark chart', 'trend line'],
      'slicer': ['slicer', 'filter slicer', 'timeline slicer'],
      'validation': ['data validation', 'cell validation', 'input validation'],
      'hyperlink': ['hyperlink', 'link', 'url', 'web link'],
      'macro': ['macro', 'vba', 'automation script', 'recorded macro'],
      // New high-priority targets
      'row': ['row', 'rows'],
      'column': ['column', 'columns', 'col', 'cols'],
      'cell': ['cell', 'cells'],
      'style': ['style', 'cell style', 'formatting style', 'custom style'],
      'connection': ['connection', 'data source', 'external data', 'data connection', 'external connection'],
      'relationship': ['relationship', 'data model', 'relation', 'link tables', 'table relationship'],
      'group': ['group', 'outline', 'grouped rows', 'grouped columns'],
      'view': ['view', 'custom view', 'saved view'],
      'scenario': ['scenario', 'what-if', 'analysis scenario', 'what if'],
      'goal': ['goal seek', 'goal', 'target value'],
      'print': ['print', 'printing', 'print area', 'print settings'],
      // Medium-priority targets
      'page': ['page', 'page setup', 'page break', 'orientation', 'paper size'],
      'header': ['header', 'page header'],
      'footer': ['footer', 'page footer'],
      'outline': ['outline', 'data outline', 'grouping structure'],
      'permission': ['permission', 'access', 'protect', 'security', 'lock', 'password'],
      'audit': ['audit', 'track changes', 'history', 'revision', 'change tracking']
    };

    // Merge with Russian patterns if needed
    if (this.currentLocale === 'ru' || this.containsRussian(text)) {
      for (const target of Object.keys(targetPatterns) as CommandTarget[]) {
        targetPatterns[target] = [...targetPatterns[target], ...this.russianTargetPatterns[target]];
      }
    }

    for (const [target, patterns] of Object.entries(targetPatterns)) {
      for (const pattern of patterns) {
        if (text.includes(pattern)) {
          return target as CommandTarget;
        }
      }
    }

    return 'range'; // Default target
  }

  /**
   * Extract parameters based on intent and target
   */
  private extractParameters(
    text: string,
    intent: CommandIntent,
    target: CommandTarget,
    context?: NLContext
  ): Record<string, any> {
    const parameters: Record<string, any> = {};

    // Extract range references
    const rangeMatch = text.match(/([a-z]+\d+(?::[a-z]+\d+)?)/gi);
    if (rangeMatch) {
      parameters.range = rangeMatch[0];
    }

    // Extract table names
    const tableMatch = text.match(/(?:table|from|in)\s+['"]?([\w\s]+?)['"]?(?:\s|$|,|with)/i);
    if (tableMatch) {
      parameters.tableName = tableMatch[1].trim();
    }

    // Extract column names
    const columnMatches = text.match(/column\s+['"]?([\w\s]+?)['"]?/gi);
    if (columnMatches) {
      parameters.columns = columnMatches.map(m =>
        m.replace(/column\s+['"]?/i, '').replace(/['"]?$/, '').trim()
      );
    }

    // Extract values/numbers
    const numberMatches = text.match(/\d+(?:\.\d+)?/g);
    if (numberMatches) {
      parameters.values = numberMatches.map(n => parseFloat(n));
    }

    // Target-specific parameter extraction
    switch (target) {
      case 'pivot':
        this.extractPivotParameters(text, parameters);
        break;
      case 'chart':
        this.extractChartParameters(text, parameters);
        break;
      case 'query':
        this.extractQueryParameters(text, parameters);
        break;
      case 'measure':
        this.extractMeasureParameters(text, parameters);
        break;
      case 'table':
        this.extractTableParameters(text, parameters);
        break;
      case 'sparkline':
        this.extractSparklineParameters(text, parameters);
        break;
      case 'slicer':
        this.extractSlicerParameters(text, parameters);
        break;
      case 'namedRange':
        this.extractNamedRangeParameters(text, parameters);
        break;
      case 'comment':
        this.extractCommentParameters(text, parameters);
        break;
      case 'row':
      case 'column':
        this.extractRowColumnParameters(text, parameters);
        break;
      case 'cell':
        this.extractCellParameters(text, parameters);
        break;
      case 'style':
        this.extractStyleParameters(text, parameters);
        break;
      case 'connection':
        this.extractConnectionParameters(text, parameters);
        break;
      case 'relationship':
        this.extractRelationshipParameters(text, parameters);
        break;
      case 'scenario':
        this.extractScenarioParameters(text, parameters);
        break;
      case 'goal':
        this.extractGoalParameters(text, parameters);
        break;
      case 'print':
        this.extractPrintParameters(text, parameters);
        break;
    }

    // Intent-specific parameter extraction
    switch (intent) {
      case 'filter':
        this.extractFilterParameters(text, parameters);
        break;
      case 'sort':
        this.extractSortParameters(text, parameters);
        break;
      case 'format':
        this.extractFormatParameters(text, parameters);
        break;
      case 'compare':
        this.extractCompareParameters(text, parameters);
        break;
      case 'merge':
        this.extractMergeParameters(text, parameters);
        break;
      case 'validate':
        this.extractValidationParameters(text, parameters);
        break;
      case 'duplicate':
        this.extractDuplicateParameters(text, parameters);
        break;
      case 'move':
        this.extractMoveParameters(text, parameters);
        break;
      case 'hide':
      case 'show':
        this.extractHideShowParameters(text, parameters);
        break;
      case 'freeze':
        this.extractFreezeParameters(text, parameters);
        break;
      case 'find':
        this.extractFindParameters(text, parameters);
        break;
      case 'replace':
        this.extractReplaceParameters(text, parameters);
        break;
      case 'group':
      case 'ungroup':
        this.extractGroupParameters(text, parameters);
        break;
      case 'convert':
        this.extractConvertParameters(text, parameters);
        break;
      case 'link':
        this.extractLinkParameters(text, parameters);
        break;
      case 'optimize':
        this.extractOptimizeParameters(text, parameters);
        break;
    }

    return parameters;
  }

  /**
   * Extract pivot table specific parameters
   */
  private extractPivotParameters(text: string, parameters: Record<string, any>): void {
    // Row fields
    const rowMatch = text.match(/(\w+)\s+in\s+rows?/gi);
    if (rowMatch) {
      parameters.rowFields = rowMatch.map(m =>
        m.replace(/\s+in\s+rows?/i, '').trim()
      );
    }

    // Column fields
    const colMatch = text.match(/(\w+)\s+in\s+columns?/gi);
    if (colMatch) {
      parameters.columnFields = colMatch.map(m =>
        m.replace(/\s+in\s+columns?/i, '').trim()
      );
    }

    // Data/Values fields
    const valueMatch = text.match(/(sum|count|average|max|min)?\s*of?\s*(\w+)\s+(?:in\s+)?(?:values?|data)/gi);
    if (valueMatch) {
      parameters.dataFields = valueMatch.map(m => {
        const aggMatch = m.match(/(sum|count|average|max|min)/i);
        const fieldMatch = m.match(/of\s+(\w+)/i) || m.match(/(\w+)\s+(?:in|as)/);
        return {
          field: fieldMatch ? fieldMatch[1] : m.replace(/\s+(?:in\s+)?(?:values?|data).*/i, '').trim(),
          aggregation: aggMatch ? aggMatch[1].toLowerCase() : 'sum'
        };
      });
    }

    // Filter fields
    const filterMatch = text.match(/filter\s+(?:by\s+)?(\w+)/gi);
    if (filterMatch) {
      parameters.filterFields = filterMatch.map(m =>
        m.replace(/filter\s+(?:by\s+)?/i, '').trim()
      );
    }
  }

  /**
   * Extract chart specific parameters
   */
  private extractChartParameters(text: string, parameters: Record<string, any>): void {
    // Chart type
    const chartTypes: Record<string, string> = {
      'column': 'columnClustered',
      'bar': 'barClustered',
      'line': 'line',
      'pie': 'pie',
      'scatter': 'xyscatter',
      'area': 'area',
      'waterfall': 'waterfall',
      'funnel': 'funnel',
      'treemap': 'treemap',
      'sunburst': 'sunburst',
      'histogram': 'histogram',
      'box': 'boxWhisker',
      'combo': 'combo'
    };

    for (const [type, value] of Object.entries(chartTypes)) {
      if (text.includes(type)) {
        parameters.chartType = value;
        break;
      }
    }

    // Title
    const titleMatch = text.match(/(?:titled?|called|named)\s+['"]([^'"]+)['"]/i) ||
                       text.match(/(?:titled?|called|named)\s+(\w+(?:\s+\w+)*)/i);
    if (titleMatch) {
      parameters.title = titleMatch[1];
    }

    // Axis information
    if (text.includes('x-axis') || text.includes('horizontal')) {
      parameters.xAxis = true;
    }
    if (text.includes('y-axis') || text.includes('vertical')) {
      parameters.yAxis = true;
    }
  }

  /**
   * Extract Power Query specific parameters
   */
  private extractQueryParameters(text: string, parameters: Record<string, any>): void {
    // Data source
    if (text.includes('excel') || text.includes('workbook')) {
      parameters.sourceType = 'excel';
    } else if (text.includes('csv')) {
      parameters.sourceType = 'csv';
    } else if (text.includes('sql') || text.includes('database')) {
      parameters.sourceType = 'sql';
    } else if (text.includes('web') || text.includes('api')) {
      parameters.sourceType = 'web';
    }

    // Transformations
    const transformations: string[] = [];
    if (text.includes('filter')) transformations.push('filter');
    if (text.includes('sort')) transformations.push('sort');
    if (text.includes('group')) transformations.push('group');
    if (text.includes('merge')) transformations.push('merge');
    if (text.includes('append')) transformations.push('append');
    if (text.includes('pivot')) transformations.push('pivot');
    if (text.includes('unpivot')) transformations.push('unpivot');
    parameters.transformations = transformations;
  }

  /**
   * Extract DAX measure specific parameters
   */
  private extractMeasureParameters(text: string, parameters: Record<string, any>): void {
    // Measure name
    const nameMatch = text.match(/(?:called|named)\s+['"]?([^'"]+?)['"]?(?:\s+as\s+|\s*=\s*)/i) ||
                      text.match(/(?:create|add)\s+(?:a\s+)?(?:measure\s+)?['"]?([^'"]+?)['"]?(?:\s+as\s+|$)/i);
    if (nameMatch) {
      parameters.measureName = nameMatch[1].trim();
    }

    // Aggregation type
    const aggMatch = text.match(/(sum|count|average|max|min|distinct\s*count)/i);
    if (aggMatch) {
      parameters.aggregation = aggMatch[1].toLowerCase().replace(/\s/g, '');
    }

    // Column reference
    const colMatch = text.match(/(?:of|from)\s+(?:the\s+)?(?:column\s+)?(\w+)/i);
    if (colMatch) {
      parameters.column = colMatch[1];
    }
  }

  /**
   * Extract table specific parameters
   */
  private extractTableParameters(text: string, parameters: Record<string, any>): void {
    // Has headers
    if (text.includes('with headers') || text.includes('has headers')) {
      parameters.hasHeaders = true;
    }

    // Table style
    const styleMatch = text.match(/style\s+(\w+)/i);
    if (styleMatch) {
      parameters.style = styleMatch[1];
    }

    // Totals row
    if (text.includes('totals') || text.includes('total row')) {
      parameters.showTotals = true;
    }
  }

  /**
   * Extract filter specific parameters
   */
  private extractFilterParameters(text: string, parameters: Record<string, any>): void {
    // Filter conditions
    const conditions: Array<{ column: string; operator: string; value: string }> = [];

    // Pattern: column = value, column > value, etc.
    const conditionPattern = /(\w+)\s*(=|<>|>|>=|<|<=)\s*['"]?([^'"\s]+)['"]?/gi;
    let match;
    while ((match = conditionPattern.exec(text)) !== null) {
      conditions.push({
        column: match[1],
        operator: match[2],
        value: match[3]
      });
    }

    if (conditions.length > 0) {
      parameters.filterConditions = conditions;
    }

    // Top N
    const topMatch = text.match(/top\s+(\d+)/i);
    if (topMatch) {
      parameters.topN = parseInt(topMatch[1]);
    }
  }

  /**
   * Extract sort specific parameters
   */
  private extractSortParameters(text: string, parameters: Record<string, any>): void {
    // Sort columns
    const sortMatches = text.match(/(?:by|sort\s+by)\s+(\w+)/gi);
    if (sortMatches) {
      parameters.sortBy = sortMatches.map(m => {
        const col = m.replace(/(?:by|sort\s+by)\s+/i, '').trim();
        const order = text.includes('descending') || text.includes('desc') || text.includes('largest')
          ? 'descending'
          : 'ascending';
        return { column: col, order };
      });
    }
  }

  /**
   * Extract format specific parameters
   */
  private extractFormatParameters(text: string, parameters: Record<string, any>): void {
    // Number format
    if (text.includes('currency') || text.includes('dollar') || text.includes('$')) {
      parameters.numberFormat = '$#,##0.00';
    } else if (text.includes('percentage') || text.includes('%')) {
      parameters.numberFormat = '0.00%';
    } else if (text.includes('date')) {
      parameters.numberFormat = 'yyyy-mm-dd';
    }

    // Colors
    const colorMatch = text.match(/(?:color|fill)\s+(\w+)/i);
    if (colorMatch) {
      parameters.fillColor = colorMatch[1];
    }

    // Font
    const sizeMatch = text.match(/size\s+(\d+)/i);
    if (sizeMatch) {
      parameters.fontSize = parseInt(sizeMatch[1]);
    }

    if (text.includes('bold')) {
      parameters.bold = true;
    }
    if (text.includes('italic')) {
      parameters.italic = true;
    }
  }

  /**
   * Extract sparkline specific parameters
   */
  private extractSparklineParameters(text: string, parameters: Record<string, any>): void {
    // Sparkline type
    const sparklineTypes: Record<string, string> = {
      'line': 'line',
      'column': 'column',
      'winloss': 'winLoss',
      'win/loss': 'winLoss'
    };

    for (const [type, value] of Object.entries(sparklineTypes)) {
      if (text.includes(type)) {
        parameters.sparklineType = value;
        break;
      }
    }

    // Data range (source data for sparkline)
    const dataRangeMatch = text.match(/(?:for|from|data)\s+([a-z]+\d+(?::[a-z]+\d+)?)/i);
    if (dataRangeMatch) {
      parameters.dataRange = dataRangeMatch[1];
    }

    // Location cell (where to place sparkline)
    const locationMatch = text.match(/(?:in|at|to)\s+(?:cell\s+)?([a-z]+\d+)/i);
    if (locationMatch) {
      parameters.locationCell = locationMatch[1];
    }

    // Style options
    if (text.includes('marker') || text.includes('markers')) {
      parameters.showMarkers = true;
    }
    if (text.includes('high point') || text.includes('highpoint')) {
      parameters.highlightHighPoint = true;
    }
    if (text.includes('low point') || text.includes('lowpoint')) {
      parameters.highlightLowPoint = true;
    }
    if (text.includes('negative')) {
      parameters.highlightNegative = true;
    }
  }

  /**
   * Extract slicer specific parameters
   */
  private extractSlicerParameters(text: string, parameters: Record<string, any>): void {
    // Slicer field (the column/field to filter)
    const fieldMatch = text.match(/(?:for|by|on)\s+(?:the\s+)?(\w+)/i) ||
                      text.match(/slicer\s+(?:for\s+)?(\w+)/i);
    if (fieldMatch) {
      parameters.slicerField = fieldMatch[1];
    }

    // Slicer type
    if (text.includes('timeline') || text.includes('date') || text.includes('time')) {
      parameters.slicerType = 'timeline';
    } else {
      parameters.slicerType = 'standard';
    }

    // Pivot table reference
    const pivotMatch = text.match(/(?:pivot|table)\s+['"]?([\w\s]+?)['"]?(?:\s|$)/i);
    if (pivotMatch) {
      parameters.pivotTable = pivotMatch[1].trim();
    }

    // Position
    const posMatch = text.match(/(?:at|position)\s+\(?(\d+)\s*,\s*(\d+)\)?/i);
    if (posMatch) {
      parameters.position = { left: parseInt(posMatch[1]), top: parseInt(posMatch[2]) };
    }
  }

  /**
   * Extract named range specific parameters
   */
  private extractNamedRangeParameters(text: string, parameters: Record<string, any>): void {
    // Range name
    const nameMatch = text.match(/(?:called|named)\s+['"]?([^'"\s]+)['"]?/i) ||
                      text.match(/(?:name|range)\s+['"]?([^'"\s]+)['"]?/i) ||
                      text.match(/['"]([^'"]+)['"]/);
    if (nameMatch) {
      parameters.rangeName = nameMatch[1].trim();
    }

    // Range address
    const addressMatch = text.match(/(?:from|for|range)\s+([a-z]+\d+(?::[a-z]+\d+)?)/i);
    if (addressMatch) {
      parameters.rangeAddress = addressMatch[1];
    }

    // Scope
    if (text.includes('workbook')) {
      parameters.scope = 'workbook';
    } else if (text.includes('worksheet') || text.includes('sheet')) {
      const sheetMatch = text.match(/(?:sheet|worksheet)\s+['"]?([^'"\s]+)['"]?/i);
      parameters.scope = sheetMatch ? sheetMatch[1] : 'worksheet';
    }

    // Comment/description
    const commentMatch = text.match(/(?:comment|description)\s+['"]([^'"]+)['"]/i);
    if (commentMatch) {
      parameters.comment = commentMatch[1];
    }
  }

  /**
   * Extract comment specific parameters
   */
  private extractCommentParameters(text: string, parameters: Record<string, any>): void {
    // Comment text
    const textMatch = text.match(/(?:saying|text|with)\s+['"]([^'"]+)['"]/i) ||
                      text.match(/(?:comment|note)\s+['"]([^'"]+)['"]/i);
    if (textMatch) {
      parameters.commentText = textMatch[1];
    }

    // Cell address
    const cellMatch = text.match(/(?:to|in|at)\s+(?:cell\s+)?([a-z]+\d+)/i);
    if (cellMatch) {
      parameters.cellAddress = cellMatch[1];
    }

    // Author
    const authorMatch = text.match(/(?:by|author)\s+['"]?([^'"\s]+)['"]?/i);
    if (authorMatch) {
      parameters.author = authorMatch[1];
    }

    // Show/hide all
    if (text.includes('show all') || text.includes('showall')) {
      parameters.showAll = true;
    }
    if (text.includes('hide all') || text.includes('hideall')) {
      parameters.hideAll = true;
    }
  }

  /**
   * Extract compare specific parameters
   */
  private extractCompareParameters(text: string, parameters: Record<string, any>): void {
    // Compare source and target
    const vsMatch = text.match(/(\w+)\s+(?:vs|versus|and|with)\s+(\w+)/i);
    if (vsMatch) {
      parameters.compareSource = vsMatch[1];
      parameters.compareTarget = vsMatch[2];
    }

    // Compare type
    if (text.includes('value')) {
      parameters.compareType = 'values';
    } else if (text.includes('structure') || text.includes('format')) {
      parameters.compareType = 'structure';
    } else if (text.includes('both')) {
      parameters.compareType = 'both';
    } else {
      parameters.compareType = 'values'; // default
    }

    // Ranges to compare
    const ranges = text.match(/([a-z]+\d+(?::[a-z]+\d+)?)/gi);
    if (ranges && ranges.length >= 2) {
      parameters.sourceRange = ranges[0];
      parameters.targetRange = ranges[1];
    }

    // Highlight differences
    if (text.includes('highlight') || text.includes('show') || text.includes('mark')) {
      parameters.highlightDifferences = true;
    }
  }

  /**
   * Extract merge specific parameters
   */
  private extractMergeParameters(text: string, parameters: Record<string, any>): void {
    // Merge type
    if (text.includes('append') || text.includes('stack')) {
      parameters.mergeType = 'append';
    } else if (text.includes('union')) {
      parameters.mergeType = 'union';
    } else if (text.includes('join') || text.includes('merge on')) {
      parameters.mergeType = 'join';
    } else {
      parameters.mergeType = 'append'; // default
    }

    // Key column for joins
    const keyMatch = text.match(/(?:on|by|key)\s+(?:column\s+)?(\w+)/i);
    if (keyMatch) {
      parameters.keyColumn = keyMatch[1];
    }

    // Source tables
    const tables = text.match(/(?:table|from)\s+['"]?([\w\s]+?)['"]?(?:\s|$|,|and|with)/gi);
    if (tables) {
      parameters.sourceTables = tables.map(t =>
        t.replace(/(?:table|from)\s+['"]?/i, '').replace(/['"]?$/, '').trim()
      );
    }

    // Target table (merge into)
    const targetMatch = text.match(/(?:into|to)\s+['"]?([\w\s]+?)['"]?(?:\s|$)/i);
    if (targetMatch) {
      parameters.targetTable = targetMatch[1].trim();
    }

    // Remove duplicates
    if (text.includes('unique') || text.includes('distinct') || text.includes('no duplicate')) {
      parameters.removeDuplicates = true;
    }
  }

  /**
   * Extract validation specific parameters
   */
  private extractValidationParameters(text: string, parameters: Record<string, any>): void {
    // Validation type
    if (text.includes('list') || text.includes('dropdown')) {
      parameters.validationType = 'list';

      // Extract list values
      const valuesMatch = text.match(/(?:values?|with)\s+(.+?)(?:\s+in\s+|\s+for\s+|$)/i);
      if (valuesMatch) {
        // Parse comma-separated values
        parameters.allowedValues = valuesMatch[1].split(/[,;]/).map(v => v.trim());
      }
    } else if (text.includes('date')) {
      parameters.validationType = 'date';
    } else if (text.includes('number') || text.includes('numeric')) {
      parameters.validationType = 'number';
    } else if (text.includes('text length') || text.includes('length')) {
      parameters.validationType = 'textLength';
    } else if (text.includes('custom') || text.includes('formula')) {
      parameters.validationType = 'custom';
    } else {
      parameters.validationType = 'any';
    }

    // Check for duplicates
    if (text.includes('duplicate')) {
      parameters.validationType = 'duplicates';
    }

    // Check for blanks
    if (text.includes('blank') || text.includes('empty')) {
      parameters.validationType = 'blanks';
    }

    // Check for errors
    if (text.includes('error')) {
      parameters.validationType = 'errors';
    }

    // Criteria for number/date validation
    const criteriaMatch = text.match(/(greater than|less than|between|equal to)\s+(\d+)/i);
    if (criteriaMatch) {
      parameters.criteria = {
        operator: criteriaMatch[1],
        value: parseFloat(criteriaMatch[2])
      };
    }

    // Range to validate
    const rangeMatch = text.match(/(?:in|for|column)\s+([a-z]+\d*(?::[a-z]+\d*)?)/i);
    if (rangeMatch) {
      parameters.validationRange = rangeMatch[1];
    }
  }

  /**
   * Extract duplicate-specific parameters
   */
  private extractDuplicateParameters(text: string, parameters: Record<string, any>): void {
    // Destination for duplicate
    const destMatch = text.match(/(?:to|into|as)\s+['"]?([\w\s]+?)['"]?(?:\s|$)/i);
    if (destMatch) {
      parameters.destination = destMatch[1].trim();
    }
    
    // Number of copies
    const countMatch = text.match(/(\d+)\s*(?:copies?|times?)/i);
    if (countMatch) {
      parameters.count = parseInt(countMatch[1]);
    }
  }

  /**
   * Extract move-specific parameters
   */
  private extractMoveParameters(text: string, parameters: Record<string, any>): void {
    // Destination
    const destMatch = text.match(/(?:to|into|before|after)\s+(?:position\s+)?(\d+|['"]?[\w\s]+?['"]?)/i);
    if (destMatch) {
      parameters.destination = destMatch[1].trim();
    }
    
    // Direction
    if (text.includes('up') || text.includes('above')) {
      parameters.direction = 'up';
    } else if (text.includes('down') || text.includes('below')) {
      parameters.direction = 'down';
    } else if (text.includes('left') || text.includes('before')) {
      parameters.direction = 'left';
    } else if (text.includes('right') || text.includes('after')) {
      parameters.direction = 'right';
    }
  }

  /**
   * Extract hide/show-specific parameters
   */
  private extractHideShowParameters(text: string, parameters: Record<string, any>): void {
    // Range of rows/columns
    const rangeMatch = text.match(/(\d+)(?:\s*(?:to|-|through)\s*(\d+))?/i);
    if (rangeMatch) {
      parameters.start = parseInt(rangeMatch[1]);
      if (rangeMatch[2]) {
        parameters.end = parseInt(rangeMatch[2]);
      }
    }
    
    // All hidden items
    if (text.includes('all')) {
      parameters.all = true;
    }
  }

  /**
   * Extract freeze-specific parameters
   */
  private extractFreezeParameters(text: string, parameters: Record<string, any>): void {
    // Freeze panes location
    const cellMatch = text.match(/(?:at|cell)\s+([a-z]+\d+)/i);
    if (cellMatch) {
      parameters.freezeCell = cellMatch[1];
    }
    
    // Freeze top row
    if (text.includes('top row') || text.includes('first row') || text.includes('header')) {
      parameters.freezeTopRow = true;
    }
    
    // Freeze first column
    if (text.includes('first column') || text.includes('left column')) {
      parameters.freezeFirstColumn = true;
    }
    
    // Unfreeze
    if (text.includes('unfreeze') || text.includes('unlock')) {
      parameters.unfreeze = true;
    }
  }

  /**
   * Extract find-specific parameters
   */
  private extractFindParameters(text: string, parameters: Record<string, any>): void {
    // Search term
    const termMatch = text.match(/(?:find|search for|locate)\s+['"]?([^'"]+?)['"]?(?:\s|$|in|within)/i);
    if (termMatch) {
      parameters.searchTerm = termMatch[1].trim();
    }
    
    // Search scope
    if (text.includes('formulas')) {
      parameters.searchIn = 'formulas';
    } else if (text.includes('values')) {
      parameters.searchIn = 'values';
    } else if (text.includes('comments')) {
      parameters.searchIn = 'comments';
    }
    
    // Match case
    if (text.includes('match case') || text.includes('case sensitive')) {
      parameters.matchCase = true;
    }
  }

  /**
   * Extract replace-specific parameters
   */
  private extractReplaceParameters(text: string, parameters: Record<string, any>): void {
    // Find and replace pattern
    const replaceMatch = text.match(/(?:replace|substitute|change)\s+['"]?([^'"]+?)['"]?\s+(?:with|to)\s+['"]?([^'"]+?)['"]?/i);
    if (replaceMatch) {
      parameters.find = replaceMatch[1].trim();
      parameters.replace = replaceMatch[2].trim();
    }
    
    // Replace all
    if (text.includes('all') || text.includes('every')) {
      parameters.replaceAll = true;
    }
  }

  /**
   * Extract group-specific parameters
   */
  private extractGroupParameters(text: string, parameters: Record<string, any>): void {
    // Group range
    const rangeMatch = text.match(/(\d+)(?:\s*(?:to|-|through)\s*(\d+))?/i);
    if (rangeMatch) {
      parameters.start = parseInt(rangeMatch[1]);
      if (rangeMatch[2]) {
        parameters.end = parseInt(rangeMatch[2]);
      }
    }
    
    // Group level
    const levelMatch = text.match(/level\s+(\d+)/i);
    if (levelMatch) {
      parameters.level = parseInt(levelMatch[1]);
    }
    
    // Collapse/expand
    if (text.includes('collapse')) {
      parameters.collapse = true;
    } else if (text.includes('expand')) {
      parameters.expand = true;
    }
  }

  /**
   * Extract convert-specific parameters
   */
  private extractConvertParameters(text: string, parameters: Record<string, any>): void {
    // Convert from/to types
    if (text.includes('to range') || text.includes('to normal')) {
      parameters.convertTo = 'range';
    } else if (text.includes('to table')) {
      parameters.convertTo = 'table';
    } else if (text.includes('to values') || text.includes('formulas to values')) {
      parameters.convertTo = 'values';
    } else if (text.includes('to formulas')) {
      parameters.convertTo = 'formulas';
    } else if (text.includes('to number') || text.includes('to numeric')) {
      parameters.convertTo = 'number';
    } else if (text.includes('to text') || text.includes('to string')) {
      parameters.convertTo = 'text';
    } else if (text.includes('to date')) {
      parameters.convertTo = 'date';
    }
  }

  /**
   * Extract link-specific parameters
   */
  private extractLinkParameters(text: string, parameters: Record<string, any>): void {
    // Link destination
    const linkMatch = text.match(/(?:link|connect|reference)\s+(?:to\s+)?['"]?([^'"]+?)['"]?/i);
    if (linkMatch) {
      parameters.destination = linkMatch[1].trim();
    }
    
    // External workbook
    if (text.includes('.xlsx') || text.includes('.xls')) {
      const fileMatch = text.match(/([\w\s]+\.xlsx?)/i);
      if (fileMatch) {
        parameters.externalFile = fileMatch[1];
      }
    }
    
    // Break links
    if (text.includes('break') || text.includes('remove')) {
      parameters.breakLink = true;
    }
    
    // Update links
    if (text.includes('update')) {
      parameters.updateLink = true;
    }
  }

  /**
   * Extract optimize-specific parameters
   */
  private extractOptimizeParameters(text: string, parameters: Record<string, any>): void {
    // Optimization type
    if (text.includes('performance') || text.includes('speed')) {
      parameters.optimizeType = 'performance';
    } else if (text.includes('size') || text.includes('compress')) {
      parameters.optimizeType = 'size';
    } else if (text.includes('formatting') || text.includes('styles')) {
      parameters.optimizeType = 'formatting';
    } else if (text.includes('images')) {
      parameters.optimizeType = 'images';
    }
  }

  /**
   * Extract row/column-specific parameters
   */
  private extractRowColumnParameters(text: string, parameters: Record<string, any>): void {
    // Row/column number or range
    const numMatch = text.match(/(\d+)(?:\s*(?:to|-|through)\s*(\d+))?/i);
    if (numMatch) {
      parameters.start = parseInt(numMatch[1]);
      if (numMatch[2]) {
        parameters.end = parseInt(numMatch[2]);
      }
    }
    
    // Column letter
    const colMatch = text.match(/(?:column|col)\s+([a-z]+)/i);
    if (colMatch) {
      parameters.columnLetter = colMatch[1].toUpperCase();
    }
    
    // Width/height
    const sizeMatch = text.match(/(?:width|height|size)\s+(?:of\s+)?(\d+)/i);
    if (sizeMatch) {
      parameters.size = parseInt(sizeMatch[1]);
    }
  }

  /**
   * Extract cell-specific parameters
   */
  private extractCellParameters(text: string, parameters: Record<string, any>): void {
    // Cell address
    const cellMatch = text.match(/([a-z]+\d+)/i);
    if (cellMatch) {
      parameters.cellAddress = cellMatch[1];
    }
  }

  /**
   * Extract style-specific parameters
   */
  private extractStyleParameters(text: string, parameters: Record<string, any>): void {
    // Style name
    const styleMatch = text.match(/(?:style|format)\s+['"]?([\w\s]+?)['"]?/i);
    if (styleMatch) {
      parameters.styleName = styleMatch[1].trim();
    }
    
    // Create custom style
    if (text.includes('create') || text.includes('new') || text.includes('custom')) {
      parameters.createStyle = true;
    }
  }

  /**
   * Extract connection-specific parameters
   */
  private extractConnectionParameters(text: string, parameters: Record<string, any>): void {
    // Connection type
    if (text.includes('sql') || text.includes('database')) {
      parameters.connectionType = 'sql';
    } else if (text.includes('web') || text.includes('url')) {
      parameters.connectionType = 'web';
    } else if (text.includes('csv') || text.includes('text file')) {
      parameters.connectionType = 'csv';
    }
    
    // Connection name
    const nameMatch = text.match(/(?:connection|source)\s+['"]?([\w\s]+?)['"]?/i);
    if (nameMatch) {
      parameters.connectionName = nameMatch[1].trim();
    }
  }

  /**
   * Extract relationship-specific parameters
   */
  private extractRelationshipParameters(text: string, parameters: Record<string, any>): void {
    // Table names
    const tablesMatch = text.match(/(?:between|from|to)\s+['"]?([\w\s]+?)['"]?\s+(?:and|to)\s+['"]?([\w\s]+?)['"]?/i);
    if (tablesMatch) {
      parameters.table1 = tablesMatch[1].trim();
      parameters.table2 = tablesMatch[2].trim();
    }
    
    // Relationship type
    if (text.includes('one to many') || text.includes('one-to-many')) {
      parameters.relationshipType = 'one-to-many';
    } else if (text.includes('many to one') || text.includes('many-to-one')) {
      parameters.relationshipType = 'many-to-one';
    } else if (text.includes('one to one') || text.includes('one-to-one')) {
      parameters.relationshipType = 'one-to-one';
    }
  }

  /**
   * Extract scenario-specific parameters
   */
  private extractScenarioParameters(text: string, parameters: Record<string, any>): void {
    // Scenario name
    const nameMatch = text.match(/(?:scenario|what-if)\s+['"]?([\w\s]+?)['"]?/i);
    if (nameMatch) {
      parameters.scenarioName = nameMatch[1].trim();
    }
    
    // Changing cells
    const cellsMatch = text.match(/(?:changing|change)\s+(?:cells?\s+)?([a-z]+\d+(?:\s*,\s*[a-z]+\d+)*)/i);
    if (cellsMatch) {
      parameters.changingCells = cellsMatch[1].split(',').map(c => c.trim());
    }
  }

  /**
   * Extract goal seek-specific parameters
   */
  private extractGoalParameters(text: string, parameters: Record<string, any>): void {
    // Target cell
    const targetMatch = text.match(/(?:target|goal|make)\s+(?:cell\s+)?([a-z]+\d+)/i);
    if (targetMatch) {
      parameters.targetCell = targetMatch[1];
    }
    
    // Target value
    const valueMatch = text.match(/(?:equal|to|value)\s+(\d+(?:\.\d+)?)/i);
    if (valueMatch) {
      parameters.targetValue = parseFloat(valueMatch[1]);
    }
    
    // Changing cell
    const changingMatch = text.match(/(?:by|changing|change)\s+(?:cell\s+)?([a-z]+\d+)/i);
    if (changingMatch) {
      parameters.changingCell = changingMatch[1];
    }
  }

  /**
   * Extract print-specific parameters
   */
  private extractPrintParameters(text: string, parameters: Record<string, any>): void {
    // Print area
    const areaMatch = text.match(/(?:print\s+area|area)\s+([a-z]+\d+(?::[a-z]+\d+)?)/i);
    if (areaMatch) {
      parameters.printArea = areaMatch[1];
    }
    
    // Page orientation
    if (text.includes('landscape')) {
      parameters.orientation = 'landscape';
    } else if (text.includes('portrait')) {
      parameters.orientation = 'portrait';
    }
    
    // Margins
    if (text.includes('narrow') || text.includes('wide') || text.includes('normal')) {
      parameters.margins = text.includes('narrow') ? 'narrow' : text.includes('wide') ? 'wide' : 'normal';
    }
  }

  /**
   * Extract constraints from the command
   */
  private extractConstraints(text: string): string[] {
    const constraints: string[] = [];

    // Conditional constraints
    if (text.includes('if') || text.includes('when')) {
      constraints.push('conditional');
    }

    // Exclusion constraints
    if (text.includes('except') || text.includes('excluding') || text.includes('not')) {
      constraints.push('exclusion');
    }

    // Range constraints
    if (text.includes('between') || text.includes('from') || text.includes('to')) {
      constraints.push('range');
    }

    // Date constraints
    if (text.includes('after') || text.includes('before') || text.includes('since')) {
      constraints.push('date');
    }

    return constraints;
  }

  /**
   * Calculate confidence score for the parsing
   */
  private calculateConfidence(
    intent: CommandIntent,
    target: CommandTarget,
    parameters: Record<string, any>,
    text: string
  ): 'high' | 'medium' | 'low' {
    let score = 0;

    // Intent detection confidence
    if (intent !== 'explain') score += 2; // Non-default intent

    // Target detection confidence
    if (target !== 'range') score += 2; // Non-default target

    // Parameter richness
    const paramCount = Object.keys(parameters).length;
    score += paramCount;

    // Text specificity
    if (text.length > 20) score += 1;
    if (/\d/.test(text)) score += 1; // Contains numbers
    if (/[A-Z]\d+/.test(text)) score += 2; // Contains cell references

    // Convert score to confidence level
    if (score >= 6) return 'high';
    if (score >= 3) return 'medium';
    return 'low';
  }

  /**
   * Generate suggestions based on context
   */
  private generateSuggestions(
    intent: CommandIntent,
    target: CommandTarget,
    context: NLContext
  ): string[] {
    const suggestions: string[] = [];

    // Suggest operations based on selected data
    if (context.dataType === 'numeric' && intent === 'analyze') {
      suggestions.push('Consider creating a summary with SUM, AVERAGE, and COUNT');
    }

    if (context.rowCount && context.rowCount > 1000 && target === 'table') {
      suggestions.push('Large dataset detected. Consider filtering before operations.');
    }

    // Suggest chart types based on data
    if (target === 'chart' && context.dataType === 'date') {
      suggestions.push('Time-series data detected. Line chart would be appropriate.');
    }

    // Suggest pivot configurations
    if (target === 'pivot' && context.availableColumns) {
      const numericCols = context.availableColumns.filter(c =>
        c.toLowerCase().includes('amount') ||
        c.toLowerCase().includes('price') ||
        c.toLowerCase().includes('quantity')
      );
      if (numericCols.length > 0) {
        suggestions.push(`Consider using ${numericCols[0]} as your values field`);
      }
    }

    return suggestions;
  }

  /**
   * Parse multiple commands from text (if separated by conjunctions)
   */
  parseMultipleCommands(text: string, context?: NLContext): ParsedCommand[] {
    // Split by common conjunctions that indicate separate commands
    const separators = /\s+(?:and then|then|and also|plus|,\s+(?:then|next))\s+/i;
    const parts = text.split(separators);

    return parts.map(part => this.parseCommand(part.trim(), context));
  }

  /**
   * Generate example commands for a specific target with bilingual support
   */
  getExampleCommands(target: CommandTarget, locale: SupportedLocale = 'en'): string[] {
    const examples: Record<CommandTarget, Record<SupportedLocale, string[]>> = {
      'pivot': {
        'en': [
          'Create a pivot table from Sales data with Region in rows and Sum of Amount in values',
          'Add Product Category to columns of the pivot table',
          'Filter the pivot to show only 2024 data'
        ],
        'ru': [
          'Создай сводную таблицу из данных Продаж с Регионом в строках и Суммой в значениях',
          'Добавь Категорию Продукта в колонки сводной таблицы',
          'Отфильтруй сводную таблицу чтобы показать только данные за 2024 год',
          'Построй сводную с Регионами по строкам и Продажами по столбцам'
        ]
      },
      'chart': {
        'en': [
          'Create a column chart of Sales by Region titled "Regional Performance"',
          'Add a trendline to the sales chart',
          'Create a combo chart with sales as columns and growth rate as a line'
        ],
        'ru': [
          'Создай столбчатую диаграмму Продаж по Регионам с названием "Региональная эффективность"',
          'Добавь линию тренда на график продаж',
          'Построй комбинированную диаграмму с продажами столбцами и темпом роста линией',
          'Сделай круговую диаграмму по долям рынка'
        ]
      },
      'table': {
        'en': [
          'Create a table from range A1:D50 with headers',
          'Apply blue table style to the current table',
          'Add a totals row to the Sales table'
        ],
        'ru': [
          'Создай таблицу из диапазона A1:D50 с заголовками',
          'Преврати диапазон в умную таблицу',
          'Сделай таблицу из выделенного диапазона с шапкой'
        ]
      },
      'query': {
        'en': [
          'Load data from the CSV file in Downloads folder',
          'Filter the query to show only active customers',
          'Merge the Orders query with Customers on CustomerID'
        ],
        'ru': [
          'Загрузи данные из CSV файла из папки Загрузки',
          'Отфильтруй запрос чтобы показать только активных клиентов',
          'Объедини запросы Заказы и Клиенты по CustomerID'
        ]
      },
      'measure': {
        'en': [
          'Create a measure Total Sales as sum of Amount from Sales table',
          'Calculate year-over-year growth as a percentage',
          'Create a running total measure for Sales'
        ],
        'ru': [
          'Создай меру Всего Продаж как сумму Суммы из таблицы Продаж',
          'Добавь вычисляемое поле Средний Чек как среднее Суммы',
          'Сделай меру Количество Уникальных Клиентов'
        ]
      },
      'range': {
        'en': [
          'Format cells A1:D10 as currency with 2 decimals',
          'Sort the selection by Sales descending',
          'Apply conditional formatting to highlight values greater than 1000'
        ],
        'ru': [
          'Отформатируй ячейки A1:D10 как валюту с 2 знаками после запятой',
          'Отсортируй выделенное по Продажам по убыванию',
          'Примени условное форматирование чтобы выделить значения больше 1000'
        ]
      },
      'worksheet': {
        'en': [
          'Create a new worksheet called "Analysis"',
          'Copy the current sheet to a new workbook',
          'Hide all sheets except the active one'
        ],
        'ru': [
          'Создай новый лист с названием "Анализ"',
          'Скопируй текущий лист в новую книгу',
          'Скрой все листы кроме активного'
        ]
      },
      'workbook': {
        'en': [
          'Protect the workbook with password',
          'Save a copy of this workbook as PDF',
          'Refresh all data connections in the workbook'
        ],
        'ru': [
          'Защити книгу паролем',
          'Сохрани копию книги как PDF',
          'Обнови все подключения к данным в книге'
        ]
      },
      'shape': {
        'en': [
          'Insert a rectangle at position (100, 100) with size 200x100',
          'Create an arrow connecting Shape1 to Shape2',
          'Add a text box with the title "Sales Report"'
        ],
        'ru': [
          'Вставь прямоугольник в позицию (100, 100) размером 200x100',
          'Создай стрелку соединяющую Фигура1 с Фигура2',
          'Добавь текстовое поле с названием "Отчет о продажах"'
        ]
      },
      'image': {
        'en': [
          'Insert the company logo at the top of the sheet',
          'Add an image from URL https://example.com/chart.png'
        ],
        'ru': [
          'Вставь логотип компании вверху листа',
          'Добавь изображение из URL https://example.com/chart.png'
        ]
      },
      'formula': {
        'en': [
          'Explain the formula in cell B5',
          'Optimize the formula in the selected range',
          'Convert this formula to use absolute references'
        ],
        'ru': [
          'Объясни формулу в ячейке B5',
          'Разбери формулу =СУММ(A1:A10)',
          'Опиши как работает формула в выделенной ячейке'
        ]
      },
      'data': {
        'en': [
          'Analyze the data in the current table',
          'Remove duplicates from column A',
          'Fill down the values in the selected range'
        ],
        'ru': [
          'Проанализируй данные в текущей таблице',
          'Удали дубликаты из столбца A',
          'Заполни вниз значения в выделенном диапазоне'
        ]
      },
      'namedRange': {
        'en': [
          'Create a named range called SalesData for A1:D100',
          'Define name TotalSales for cell E5',
          'Delete the named range OldData'
        ],
        'ru': [
          'Создай именованный диапазон ДанныеПродаж для A1:D100',
          'Задай имя ИтогоПродаж для ячейки E5',
          'Удали именованный диапазон СтарыеДанные'
        ]
      },
      'comment': {
        'en': [
          'Add a comment to cell A1 saying "Enter sales amount here"',
          'Delete all comments in the selected range',
          'Show all comments on this worksheet'
        ],
        'ru': [
          'Добавь комментарий к ячейке A1 "Введите сумму продаж"',
          'Удали все комментарии в выделенном диапазоне',
          'Покажи все комментарии на этом листе'
        ]
      },
      'sparkline': {
        'en': [
          'Create a line sparkline in column F showing sales trend',
          'Add column sparklines for the data in rows 1-10',
          'Delete sparklines in the selected cells'
        ],
        'ru': [
          'Создай линейный спарклайн в колонке F для тренда продаж',
          'Добавь столбчатые спарклайны для данных в строках 1-10',
          'Удали спарклайны в выделенных ячейках'
        ]
      },
      'slicer': {
        'en': [
          'Add a slicer for Region to the pivot table',
          'Create a timeline slicer for the Date column',
          'Delete the Category slicer'
        ],
        'ru': [
          'Добавь слайсер для Региона к сводной таблице',
          'Создай слайсер временной шкалы для колонки Дата',
          'Удали слайсер Категории'
        ]
      },
      'validation': {
        'en': [
          'Set data validation to allow only numbers in column A',
          'Add dropdown validation with Yes, No, Maybe options',
          'Create date validation requiring dates between 2020 and 2025'
        ],
        'ru': [
          'Установи проверку данных чтобы разрешить только числа в столбце A',
          'Добавь выпадающий список с опциями Да, Нет, Может быть',
          'Создай проверку дат требующую даты между 2020 и 2025'
        ]
      },
      'hyperlink': {
        'en': [
          'Add a hyperlink to cell A1 linking to Sheet2!B5',
          'Create a link to https://example.com in the selected cell',
          'Remove all hyperlinks from column C'
        ],
        'ru': [
          'Добавь ссылку в ячейку A1 на Sheet2!B5',
          'Создай ссылку на https://example.com в выделенной ячейке',
          'Удали все ссылки из столбца C'
        ]
      },
      'macro': {
        'en': [
          'Create a macro to format the selected range',
          'Explain the VBA code in Module1',
          'Record a macro for data cleanup'
        ],
        'ru': [
          'Создай макрос для форматирования выделенного диапазона',
          'Объясни VBA код в Module1',
          'Запиши макрос для очистки данных'
        ]
      },
      // New high-priority targets
      'row': {
        'en': [
          'Insert 3 rows above row 5',
          'Delete rows 10 through 15',
          'Hide rows 20 to 25',
          'Resize row 5 to height 30',
          'Group rows 5 to 10'
        ],
        'ru': [
          'Вставь 3 строки выше строки 5',
          'Удали строки с 10 по 15',
          'Скрой строки с 20 по 25',
          'Измени высоту строки 5 до 30',
          'Сгруппируй строки с 5 по 10'
        ]
      },
      'column': {
        'en': [
          'Insert a column before column B',
          'Delete column D',
          'Hide columns C through E',
          'Resize column A to width 15',
          'Move column B to position D'
        ],
        'ru': [
          'Вставь столбец перед столбцом B',
          'Удали столбец D',
          'Скрой столбцы с C по E',
          'Измени ширину столбца A до 15',
          'Перемести столбец B в позицию D'
        ]
      },
      'cell': {
        'en': [
          'Format cell A1 as currency',
          'Add comment to cell B5',
          'Link cell C10 to Sheet2!A1',
          'Clear cell D20'
        ],
        'ru': [
          'Отформатируй ячейку A1 как валюту',
          'Добавь комментарий к ячейке B5',
          'Свяжи ячейку C10 с Sheet2!A1',
          'Очисти ячейку D20'
        ]
      },
      'style': {
        'en': [
          'Apply Heading 1 style to row 1',
          'Create a custom style called Highlight',
          'Copy style from cell A1 to B1',
          'List all available styles'
        ],
        'ru': [
          'Примени стиль Заголовок 1 к строке 1',
          'Создай пользовательский стиль Выделение',
          'Скопируй стиль из ячейки A1 в B1',
          'Покажи все доступные стили'
        ]
      },
      'connection': {
        'en': [
          'Create connection to SQL Server database',
          'Refresh all data connections',
          'List all external connections',
          'Remove connection to Database1'
        ],
        'ru': [
          'Создай подключение к базе данных SQL Server',
          'Обнови все подключения к данным',
          'Покажи все внешние подключения',
          'Удали подключение к Database1'
        ]
      },
      'relationship': {
        'en': [
          'Create relationship between Sales and Customers tables',
          'Show all relationships in the data model',
          'Delete relationship between Tables A and B'
        ],
        'ru': [
          'Создай связь между таблицами Продажи и Клиенты',
          'Покажи все связи в модели данных',
          'Удали связь между таблицами A и B'
        ]
      },
      'group': {
        'en': [
          'Group rows 5 to 10',
          'Ungroup all groups',
          'Collapse group level 2',
          'Show outline symbols'
        ],
        'ru': [
          'Сгруппируй строки с 5 по 10',
          'Разгруппируй все группы',
          'Сверни группу уровня 2',
          'Покажи символы структуры'
        ]
      },
      'view': {
        'en': [
          'Create a view called Print View',
          'Switch to view Data Entry',
          'Delete view Old View',
          'List all views'
        ],
        'ru': [
          'Создай представление Печать',
          'Переключись на представление Ввод данных',
          'Удали представление Старое',
          'Покажи все представления'
        ]
      },
      'scenario': {
        'en': [
          'Create scenario Best Case',
          'Show scenario summary',
          'Switch to scenario Worst Case',
          'Merge scenarios from Budget.xlsx'
        ],
        'ru': [
          'Создай сценарий Лучший случай',
          'Покажи сводку сценариев',
          'Переключись на сценарий Худший случай',
          'Объедини сценарии из Budget.xlsx'
        ]
      },
      'goal': {
        'en': [
          'Use goal seek to make B10 equal 1000 by changing B5',
          'Find value for cell C20 to achieve target in D20'
        ],
        'ru': [
          'Используй подбор параметра чтобы сделать B10 равным 1000 изменяя B5',
          'Найди значение для ячейки C20 чтобы достичь цели в D20'
        ]
      },
      'print': {
        'en': [
          'Set print area to A1:F50',
          'Add header Monthly Report',
          'Set margins to narrow',
          'Print preview'
        ],
        'ru': [
          'Установи область печати A1:F50',
          'Добавь заголовок Ежемесячный отчет',
          'Установи поля узкие',
          'Предварительный просмотр печати'
        ]
      },
      // Medium-priority targets
      'page': {
        'en': [
          'Set page orientation to landscape',
          'Set paper size to A4',
          'Adjust page breaks',
          'Set print quality to high'
        ],
        'ru': [
          'Установи ориентацию страницы альбомная',
          'Установи размер бумаги A4',
          'Настрой разрывы страниц',
          'Установи качество печати высокое'
        ]
      },
      'header': {
        'en': [
          'Add header Company Name',
          'Insert page number in header',
          'Add date to header',
          'Remove header'
        ],
        'ru': [
          'Добавь заголовок Название компании',
          'Вставь номер страницы в заголовок',
          'Добавь дату в заголовок',
          'Удали заголовок'
        ]
      },
      'footer': {
        'en': [
          'Add footer with page numbers',
          'Insert date in footer',
          'Add custom text to footer',
          'Remove footer'
        ],
        'ru': [
          'Добавь нижний колонтитул с номерами страниц',
          'Вставь дату в нижний колонтитул',
          'Добавь пользовательский текст в нижний колонтитул',
          'Удали нижний колонтитул'
        ]
      },
      'outline': {
        'en': [
          'Create outline for this data',
          'Show outline symbols',
          'Auto outline this range',
          'Clear outline'
        ],
        'ru': [
          'Создай структуру для этих данных',
          'Покажи символы структуры',
          'Автоструктура для этого диапазона',
          'Очисти структуру'
        ]
      },
      'permission': {
        'en': [
          'Protect this sheet with password',
          'Allow users to edit range A1:B10',
          'Remove protection from Sheet2',
          'Set workbook permissions'
        ],
        'ru': [
          'Защити этот лист паролем',
          'Разреши пользователям редактировать диапазон A1:B10',
          'Сними защиту с Sheet2',
          'Установи разрешения книги'
        ]
      },
      'audit': {
        'en': [
          'Show change history',
          'Track changes in this workbook',
          'Highlight all changes',
          'Accept/reject changes'
        ],
        'ru': [
          'Покажи историю изменений',
          'Отслеживай изменения в этой книге',
          'Выдели все изменения',
          'Принять/отклонить изменения'
        ]
      }
    };
          'Create a dropdown list with values Yes, No, Maybe',
          'Remove data validation from the selected range'
        ],
        'ru': [
          'Установи проверку данных для ввода только чисел в колонку A',
          'Создай выпадающий список со значениями Да, Нет, Возможно',
          'Удали проверку данных из выделенного диапазона'
        ]
      },
      'hyperlink': {
        'en': [
          'Add a hyperlink to cell A1 linking to https://example.com',
          'Create a link to Sheet2 cell B5',
          'Remove all hyperlinks in the selected range'
        ],
        'ru': [
          'Добавь ссылку в ячейку A1 на https://example.com',
          'Создай ссылку на ячейку B5 листа Лист2',
          'Удали все ссылки из выделенного диапазона'
        ]
      },
      'macro': {
        'en': [
          'Record a macro to format the selected range',
          'Run the macro called FormatData',
          'Delete the old macro CleanupData'
        ],
        'ru': [
          'Запиши макрос для форматирования выделенного диапазона',
          'Запусти макрос ФорматироватьДанные',
          'Удали старый макрос ОчисткаДанных'
        ]
      },
      // New high-priority targets
      'row': {
        'en': [
          'Insert 3 rows above row 5',
          'Delete rows 10 through 15',
          'Hide rows 20 to 25',
          'Resize row 5 to height 30',
          'Group rows 5 to 10'
        ],
        'ru': [
          'Вставь 3 строки выше строки 5',
          'Удали строки с 10 по 15',
          'Скрой строки с 20 по 25',
          'Измени высоту строки 5 до 30',
          'Сгруппируй строки с 5 по 10'
        ]
      },
      'column': {
        'en': [
          'Insert a column before column B',
          'Delete column D',
          'Hide columns C through E',
          'Resize column A to width 15',
          'Move column B to position D'
        ],
        'ru': [
          'Вставь столбец перед столбцом B',
          'Удали столбец D',
          'Скрой столбцы с C по E',
          'Измени ширину столбца A до 15',
          'Перемести столбец B в позицию D'
        ]
      },
      'cell': {
        'en': [
          'Format cell A1 as currency',
          'Add comment to cell B5',
          'Link cell C10 to Sheet2!A1',
          'Clear cell D20'
        ],
        'ru': [
          'Отформатируй ячейку A1 как валюту',
          'Добавь комментарий к ячейке B5',
          'Свяжи ячейку C10 с Sheet2!A1',
          'Очисти ячейку D20'
        ]
      },
      'style': {
        'en': [
          'Apply Heading 1 style to row 1',
          'Create a custom style called Highlight',
          'Copy style from cell A1 to B1',
          'List all available styles'
        ],
        'ru': [
          'Примени стиль Заголовок 1 к строке 1',
          'Создай пользовательский стиль Выделение',
          'Скопируй стиль из ячейки A1 в B1',
          'Покажи все доступные стили'
        ]
      },
      'connection': {
        'en': [
          'Create connection to SQL Server database',
          'Refresh all data connections',
          'List all external connections',
          'Remove connection to Database1'
        ],
        'ru': [
          'Создай подключение к базе данных SQL Server',
          'Обнови все подключения к данным',
          'Покажи все внешние подключения',
          'Удали подключение к Database1'
        ]
      },
      'relationship': {
        'en': [
          'Create relationship between Sales and Customers tables',
          'Show all relationships in the data model',
          'Delete relationship between Tables A and B'
        ],
        'ru': [
          'Создай связь между таблицами Продажи и Клиенты',
          'Покажи все связи в модели данных',
          'Удали связь между таблицами A и B'
        ]
      },
      'group': {
        'en': [
          'Group rows 5 to 10',
          'Ungroup all groups',
          'Collapse group level 2',
          'Show outline symbols'
        ],
        'ru': [
          'Сгруппируй строки с 5 по 10',
          'Разгруппируй все группы',
          'Сверни группу уровня 2',
          'Покажи символы структуры'
        ]
      },
      'view': {
        'en': [
          'Create a view called Print View',
          'Switch to view Data Entry',
          'Delete view Old View',
          'List all views'
        ],
        'ru': [
          'Создай представление Печать',
          'Переключись на представление Ввод данных',
          'Удали представление Старое',
          'Покажи все представления'
        ]
      },
      'scenario': {
        'en': [
          'Create scenario Best Case',
          'Show scenario summary',
          'Switch to scenario Worst Case',
          'Merge scenarios from Budget.xlsx'
        ],
        'ru': [
          'Создай сценарий Лучший случай',
          'Покажи сводку сценариев',
          'Переключись на сценарий Худший случай',
          'Объедини сценарии из Budget.xlsx'
        ]
      },
      'goal': {
        'en': [
          'Use goal seek to make B10 equal 1000 by changing B5',
          'Find value for cell C20 to achieve target in D20'
        ],
        'ru': [
          'Используй подбор параметра чтобы сделать B10 равным 1000 изменяя B5',
          'Найди значение для ячейки C20 чтобы достичь цели в D20'
        ]
      },
      'print': {
        'en': [
          'Set print area to A1:F50',
          'Add header Monthly Report',
          'Set margins to narrow',
          'Print preview'
        ],
        'ru': [
          'Установи область печати A1:F50',
          'Добавь заголовок Ежемесячный отчет',
          'Установи поля узкие',
          'Предварительный просмотр печати'
        ]
      },
      // Medium-priority targets
      'page': {
        'en': [
          'Set page orientation to landscape',
          'Set paper size to A4',
          'Adjust page breaks',
          'Set print quality to high'
        ],
        'ru': [
          'Установи ориентацию страницы альбомная',
          'Установи размер бумаги A4',
          'Настрой разрывы страниц',
          'Установи качество печати высокое'
        ]
      },
      'header': {
        'en': [
          'Add header Company Name',
          'Insert page number in header',
          'Add date to header',
          'Remove header'
        ],
        'ru': [
          'Добавь заголовок Название компании',
          'Вставь номер страницы в заголовок',
          'Добавь дату в заголовок',
          'Удали заголовок'
        ]
      },
      'footer': {
        'en': [
          'Add footer with page numbers',
          'Insert date in footer',
          'Add custom text to footer',
          'Remove footer'
        ],
        'ru': [
          'Добавь нижний колонтитул с номерами страниц',
          'Вставь дату в нижний колонтитул',
          'Добавь пользовательский текст в нижний колонтитул',
          'Удали нижний колонтитул'
        ]
      },
      'outline': {
        'en': [
          'Create outline for this data',
          'Show outline symbols',
          'Auto outline this range',
          'Clear outline'
        ],
        'ru': [
          'Создай структуру для этих данных',
          'Покажи символы структуры',
          'Автоструктура для этого диапазона',
          'Очисти структуру'
        ]
      },
      'permission': {
        'en': [
          'Protect this sheet with password',
          'Allow users to edit range A1:B10',
          'Remove protection from Sheet2',
          'Set workbook permissions'
        ],
        'ru': [
          'Защити этот лист паролем',
          'Разреши пользователям редактировать диапазон A1:B10',
          'Сними защиту с Sheet2',
          'Установи разрешения книги'
        ]
      },
      'audit': {
        'en': [
          'Show change history',
          'Track changes in this workbook',
          'Highlight all changes',
          'Accept/reject changes'
        ],
        'ru': [
          'Покажи историю изменений',
          'Отслеживай изменения в этой книге',
          'Выдели все изменения',
          'Принять/отклонить изменения'
        ]
      }
    };

    return examples[target]?.[locale] || examples[target]?.['en'] || ['No examples available for this target'];
  }

  /**
   * Optimize voice commands for speech recognition
   */
  optimizeVoiceCommand(text: string): string {
    // Common speech recognition errors and corrections
    const speechCorrections: Array<{ from: RegExp; to: string }> = [
      // English corrections
      { from: /sum\s+of/gi, to: 'sum of' },
      { from: /pie\s+chart/gi, to: 'pie chart' },
      { from: /sell\s+/gi, to: 'cell ' },
      { from: /rose\s+/gi, to: 'rows ' },
      { from: /call\s+/gi, to: 'column ' },
      { from: /pivot\s+able/gi, to: 'pivot table' },
      { from: /hi\s+chart/gi, to: 'high chart' },
      // Russian corrections
      { from: /сам\s+из/gi, to: 'сумм' },
      { from: /столб цац/gi, to: 'столбцах' },
      { from: /свод ная/gi, to: 'сводная' },
      { from: /график\s+а/gi, to: 'графика' },
      { from: /формул\s+а/gi, to: 'формула' }
    ];

    let optimized = text;
    for (const correction of speechCorrections) {
      optimized = optimized.replace(correction.from, correction.to);
    }

    // Remove filler words
    optimized = optimized
      .replace(/\b(um|uh|ah|er|эм|ааа|это самое)\b/gi, '')
      .replace(/\s+/g, ' ')
      .trim();

    return optimized;
  }

  /**
   * Normalize Russian text by handling grammatical cases
   */
  normalizeRussianText(text: string): string {
    return text
      .toLowerCase()
      .replace(/строк(а|и|е|у|ой|ах)/gi, 'строка')
      .replace(/столб(ец|ца|це|цом|цах)/gi, 'столбец')
      .replace(/таблиц(а|ы|е|у|ей|ах)/gi, 'таблица')
      .replace(/диаграмм(а|ы|е|у|ой|ах)/gi, 'диаграмма')
      .replace(/график(а|у|ом|е|и|ов)/gi, 'график')
      .replace(/формул(а|ы|е|у|ой|ах)/gi, 'формула')
      .replace(/ячей(ка|ки|ке|ку|кой|ках)/gi, 'ячейка');
  }

  /**
   * Validate if a command can be executed with the current context
   */
  validateCommandForContext(command: ParsedCommand, context: NLContext): {
    valid: boolean;
    missingRequirements: string[];
  } {
    const missingRequirements: string[] = [];

    // Check if selection is required
    if (!context.selectedRange) {
      if (command.target === 'range' || command.target === 'table') {
        missingRequirements.push('A cell range must be selected');
      }
    }

    // Check if table is required
    if (command.target === 'pivot' && !context.activeTable && !context.selectedRange) {
      missingRequirements.push('A data source (table or range) must be specified');
    }

    // Check data type compatibility
    if (command.intent === 'calculate' && context.dataType === 'text') {
      missingRequirements.push('Numeric data is required for calculations');
    }

    return {
      valid: missingRequirements.length === 0,
      missingRequirements
    };
  }
}

export default NaturalLanguageCommandParser.getInstance();
