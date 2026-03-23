// Diagram Service - Shapes, SmartArt, and Basic Diagram Operations
// Phase 5 Implementation - SmartArt & Shapes Support

export type ShapeType =
  | 'rectangle'
  | 'roundedRectangle'
  | 'ellipse'
  | 'circle'
  | 'triangle'
  | 'diamond'
  | 'arrow'
  | 'line'
  | 'curvedLine'
  | 'star'
  | 'pentagon'
  | 'hexagon'
  | 'cloud'
  | 'heart'
  | 'cross';

export type ConnectorType =
  | 'straight'
  | 'elbow'
  | 'curved';

export interface ShapeConfig {
  type: ShapeType;
  left: number;
  top: number;
  width: number;
  height: number;
  name?: string;
}

export interface ShapeFormatting {
  fill?: {
    color?: string;
    transparency?: number;
    gradient?: {
      type: 'linear' | 'radial';
      color1: string;
      color2: string;
      angle?: number;
    };
  };
  outline?: {
    color?: string;
    weight?: number;
    style?: 'solid' | 'dash' | 'dot' | 'dashDot';
  };
  effects?: {
    shadow?: {
      enabled: boolean;
      blur?: number;
      offset?: { x: number; y: number };
      color?: string;
    };
    glow?: {
      enabled: boolean;
      color?: string;
      size?: number;
    };
    reflection?: {
      enabled: boolean;
      transparency?: number;
      size?: number;
    };
  };
  text?: {
    content?: string;
    font?: {
      name?: string;
      size?: number;
      bold?: boolean;
      italic?: boolean;
      color?: string;
    };
    alignment?: {
      horizontal?: 'left' | 'center' | 'right';
      vertical?: 'top' | 'middle' | 'bottom';
    };
  };
}

export interface ConnectorConfig {
  startShape: string;  // Shape name or ID
  endShape: string;    // Shape name or ID
  type: ConnectorType;
  name?: string;
}

export interface ConnectorFormatting {
  color?: string;
  weight?: number;
  style?: 'solid' | 'dash' | 'dot';
  beginArrowhead?: 'none' | 'arrow' | 'diamond' | 'oval';
  endArrowhead?: 'none' | 'arrow' | 'diamond' | 'oval';
}

export interface SmartArtConfig {
  type: 'list' | 'process' | 'cycle' | 'hierarchy' | 'relationship' | 'matrix' | 'pyramid';
  data: Array<{
    text: string;
    level?: number;
    children?: Array<{ text: string }>;
  }>;
  left: number;
  top: number;
  width: number;
  height: number;
}

export class DiagramService {
  private static instance: DiagramService;

  private constructor() {}

  static getInstance(): DiagramService {
    if (!DiagramService.instance) {
      DiagramService.instance = new DiagramService();
    }
    return DiagramService.instance;
  }

  // ============================================================================
  // SHAPE OPERATIONS
  // ============================================================================

  /**
   * Insert a shape into the worksheet
   */
  async insertShape(
    config: ShapeConfig,
    formatting?: ShapeFormatting,
    worksheetName?: string
  ): Promise<string> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      // Map shape type to Excel shape type
      const shapeType = this.mapShapeType(config.type);

      // Create shape
      const shape = worksheet.shapes.addGeometricShape(
        shapeType,
        {
          left: config.left,
          top: config.top,
          height: config.height,
          width: config.width
        }
      );

      if (config.name) {
        shape.name = config.name;
      }

      // Apply formatting
      if (formatting) {
        await this.applyShapeFormatting(context, shape, formatting);
      }

      shape.load('name');
      await context.sync();

      return shape.name;
    });
  }

  /**
   * Insert multiple shapes at once
   */
  async insertShapes(
    configs: Array<{ config: ShapeConfig; formatting?: ShapeFormatting }>,
    worksheetName?: string
  ): Promise<string[]> {
    const names: string[] = [];

    for (const { config, formatting } of configs) {
      const name = await this.insertShape(config, formatting, worksheetName);
      names.push(name);
    }

    return names;
  }

  /**
   * Delete a shape
   */
  async deleteShape(shapeName: string, worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.getActiveWorksheet();

      const shape = worksheet.shapes.getItem(shapeName);
      shape.delete();

      await context.sync();
    });
  }

  /**
   * Move a shape to a new position
   */
  async moveShape(
    shapeName: string,
    left: number,
    top: number,
    worksheetName?: string
  ): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.getActiveWorksheet();

      const shape = worksheet.shapes.getItem(shapeName);
      shape.left = left;
      shape.top = top;

      await context.sync();
    });
  }

  /**
   * Resize a shape
   */
  async resizeShape(
    shapeName: string,
    width: number,
    height: number,
    worksheetName?: string
  ): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.getActiveWorksheet();

      const shape = worksheet.shapes.getItem(shapeName);
      shape.width = width;
      shape.height = height;

      await context.sync();
    });
  }

  // ============================================================================
  // CONNECTOR OPERATIONS
  // ============================================================================

  /**
   * Create a connector between two shapes
   */
  async createConnector(
    config: ConnectorConfig,
    formatting?: ConnectorFormatting,
    worksheetName?: string
  ): Promise<string> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.getActiveWorksheet();

      // Get the shapes to connect
      const startShape = worksheet.shapes.getItem(config.startShape);
      const endShape = worksheet.shapes.getItem(config.endShape);

      startShape.load('id');
      endShape.load('id');
      await context.sync();

      // Create connector
      const connector = worksheet.shapes.addConnector(
        this.mapConnectorType(config.type),
        0, 0, 0, 0 // Initial position, will be attached
      );

      if (config.name) {
        connector.name = config.name;
      }

      // Connect to shapes
      connector.connectBeginShape(startShape, Excel.ConnectorSite.auto);
      connector.connectEndShape(endShape, Excel.ConnectorSite.auto);

      // Apply formatting
      if (formatting) {
        await this.applyConnectorFormatting(context, connector, formatting);
      }

      connector.load('name');
      await context.sync();

      return connector.name;
    });
  }

  /**
   * Update connector formatting
   */
  async formatConnector(
    connectorName: string,
    formatting: ConnectorFormatting,
    worksheetName?: string
  ): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.getActiveWorksheet();

      const connector = worksheet.shapes.getItem(connectorName);
      await this.applyConnectorFormatting(context, connector, formatting);

      await context.sync();
    });
  }

  // ============================================================================
  // SMARTART OPERATIONS (Limited Support)
  // ============================================================================

  /**
   * Insert SmartArt graphic (basic implementation)
   * Note: Full SmartArt support requires Office API capabilities
   */
  async insertSmartArt(
    config: SmartArtConfig,
    worksheetName?: string
  ): Promise<string> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.getActiveWorksheet();

      // Note: Full SmartArt insertion may not be available in Office.js
      // This is a placeholder that creates a basic diagram using shapes

      // For now, create a simple representation using shapes and text
      const shape = worksheet.shapes.addGeometricShape(
        Excel.GeometricShapeType.rectangle,
        {
          left: config.left,
          top: config.top,
          height: config.height,
          width: config.width
        }
      );

      shape.name = `SmartArt_${config.type}_${Date.now()}`;

      // Add text
      if (config.data.length > 0) {
        shape.textFrame.textRange.text = config.data.map(item => item.text).join('\n');
      }

      shape.load('name');
      await context.sync();

      return shape.name;
    });
  }

  // ============================================================================
  // IMAGE OPERATIONS
  // ============================================================================

  /**
   * Insert an image from URL
   */
  async insertImageFromUrl(
    url: string,
    left: number,
    top: number,
    width?: number,
    height?: number,
    worksheetName?: string
  ): Promise<string> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.getActiveWorksheet();

      const image = worksheet.shapes.addImage(
        url,
        {
          left,
          top,
          width: width || 200,
          height: height || 200
        }
      );

      image.name = `Image_${Date.now()}`;
      image.load('name');
      await context.sync();

      return image.name;
    });
  }

  /**
   * Insert an image from base64
   */
  async insertImageFromBase64(
    base64: string,
    left: number,
    top: number,
    width?: number,
    height?: number,
    worksheetName?: string
  ): Promise<string> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.getActiveWorksheet();

      const image = worksheet.shapes.addImage(
        base64,
        {
          left,
          top,
          width: width || 200,
          height: height || 200
        }
      );

      image.name = `Image_${Date.now()}`;
      image.load('name');
      await context.sync();

      return image.name;
    });
  }

  // ============================================================================
  // HELPER METHODS
  // ============================================================================

  private mapShapeType(type: ShapeType): Excel.GeometricShapeType {
    const mapping: Record<ShapeType, Excel.GeometricShapeType> = {
      'rectangle': Excel.GeometricShapeType.rectangle,
      'roundedRectangle': Excel.GeometricShapeType.roundedRectangle,
      'ellipse': Excel.GeometricShapeType.ellipse,
      'circle': Excel.GeometricShapeType.oval,
      'triangle': Excel.GeometricShapeType.triangle,
      'diamond': Excel.GeometricShapeType.diamond,
      'arrow': Excel.GeometricShapeType.rightArrow,
      'line': Excel.GeometricShapeType.line,
      'curvedLine': Excel.GeometricShapeType.curvedConnector3,
      'star': Excel.GeometricShapeType.star5,
      'pentagon': Excel.GeometricShapeType.pentagon,
      'hexagon': Excel.GeometricShapeType.hexagon,
      'cloud': Excel.GeometricShapeType.cloud,
      'heart': Excel.GeometricShapeType.heart,
      'cross': Excel.GeometricShapeType.plus
    };

    return mapping[type] || Excel.GeometricShapeType.rectangle;
  }

  private mapConnectorType(type: ConnectorType): Excel.ConnectorType {
    const mapping: Record<ConnectorType, Excel.ConnectorType> = {
      'straight': Excel.ConnectorType.straight,
      'elbow': Excel.ConnectorType.elbow,
      'curved': Excel.ConnectorType.curve
    };

    return mapping[type] || Excel.ConnectorType.straight;
  }

  private async applyShapeFormatting(
    context: Excel.RequestContext,
    shape: Excel.Shape,
    formatting: ShapeFormatting
  ): Promise<void> {
    // Fill formatting
    if (formatting.fill) {
      if (formatting.fill.color) {
        shape.fill.setSolidColor(formatting.fill.color);
      }
      if (formatting.fill.transparency !== undefined) {
        shape.fill.transparency = formatting.fill.transparency;
      }
    }

    // Outline formatting
    if (formatting.outline) {
      if (formatting.outline.color) {
        shape.lineFormat.color = formatting.outline.color;
      }
      if (formatting.outline.weight) {
        shape.lineFormat.weight = formatting.outline.weight;
      }
      if (formatting.outline.style) {
        const dashStyleMap: Record<string, Excel.ShapeLineDashStyle> = {
          'solid': Excel.ShapeLineDashStyle.solid,
          'dash': Excel.ShapeLineDashStyle.dash,
          'dot': Excel.ShapeLineDashStyle.dot,
          'dashDot': Excel.ShapeLineDashStyle.dashDot
        };
        shape.lineFormat.dashStyle = dashStyleMap[formatting.outline.style] || Excel.ShapeLineDashStyle.solid;
      }
    }

    // Text formatting
    if (formatting.text) {
      if (formatting.text.content) {
        shape.textFrame.textRange.text = formatting.text.content;
      }
      if (formatting.text.font) {
        const font = shape.textFrame.textRange.font;
        if (formatting.text.font.name) font.name = formatting.text.font.name;
        if (formatting.text.font.size) font.size = formatting.text.font.size;
        if (formatting.text.font.bold) font.bold = formatting.text.font.bold;
        if (formatting.text.font.italic) font.italic = formatting.text.font.italic;
        if (formatting.text.font.color) font.color = formatting.text.font.color;
      }
    }
  }

  private async applyConnectorFormatting(
    context: Excel.RequestContext,
    connector: Excel.Shape,
    formatting: ConnectorFormatting
  ): Promise<void> {
    if (formatting.color) {
      connector.lineFormat.color = formatting.color;
    }
    if (formatting.weight) {
      connector.lineFormat.weight = formatting.weight;
    }
    if (formatting.style) {
      const dashStyleMap: Record<string, Excel.ShapeLineDashStyle> = {
        'solid': Excel.ShapeLineDashStyle.solid,
        'dash': Excel.ShapeLineDashStyle.dash,
        'dot': Excel.ShapeLineDashStyle.dot
      };
      connector.lineFormat.dashStyle = dashStyleMap[formatting.style] || Excel.ShapeLineDashStyle.solid;
    }
    if (formatting.beginArrowhead) {
      connector.lineFormat.beginArrowheadStyle = this.mapArrowhead(formatting.beginArrowhead);
    }
    if (formatting.endArrowhead) {
      connector.lineFormat.endArrowheadStyle = this.mapArrowhead(formatting.endArrowhead);
    }
  }

  private mapArrowhead(style: string): Excel.ArrowheadStyle {
    const mapping: Record<string, Excel.ArrowheadStyle> = {
      'none': Excel.ArrowheadStyle.none,
      'arrow': Excel.ArrowheadStyle.triangle,
      'diamond': Excel.ArrowheadStyle.diamond,
      'oval': Excel.ArrowheadStyle.oval
    };
    return mapping[style] || Excel.ArrowheadStyle.none;
  }

  // ============================================================================
  // UTILITY METHODS
  // ============================================================================

  /**
   * Get all shapes in a worksheet
   */
  async getAllShapes(worksheetName?: string): Promise<Array<{ name: string; type: string }>> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.getActiveWorksheet();

      const shapes = worksheet.shapes;
      shapes.load('items/name, items/type');
      await context.sync();

      return shapes.items.map(shape => ({
        name: shape.name,
        type: shape.type as string
      }));
    });
  }

  /**
   * Group multiple shapes
   * @throws {Error} Shape grouping is not supported in this Office.js version
   */
  async groupShapes(shapeNames: string[], worksheetName?: string): Promise<string> {
    throw new Error(
      'Shape grouping is not supported in this Office.js version. ' +
      'Please use Excel desktop to group shapes manually.'
    );
  }

  /**
   * Ungroup a shape group
   * @throws {Error} Shape ungrouping is not supported in this Office.js version
   */
  async ungroupShapes(groupName: string, worksheetName?: string): Promise<void> {
    throw new Error(
      'Shape ungrouping is not supported in this Office.js version. ' +
      'Please use Excel desktop to ungroup shapes manually.'
    );
  }
}

export default DiagramService.getInstance();
