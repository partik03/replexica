import { 
  Document, 
  Packer, 
  Paragraph, 
  TextRun, 
  Table, 
  TableRow, 
  TableCell,
  HeadingLevel,
  ImageRun,
  IParagraphOptions
} from 'docx';
import { ILoader } from './_types';
import { createLoader } from './_utils';

interface DocElementPath {
  type: 'paragraph' | 'heading' | 'table-cell' | 'image-caption';
  indexes: number[];
}

function parsePathString(path: string): DocElementPath {
  const [type, ...indexStrs] = path.split('/');
  return {
    type: type as DocElementPath['type'],
    indexes: indexStrs.map(i => parseInt(i))
  };
}

function createPathString(path: DocElementPath): string {
  return `${path.type}/${path.indexes.join('/')}`;
}

function normalizeTextContent(text: string): string {
  const trimmed = text.trim();
  if (!trimmed) return '';
  return trimmed;
}

function getHeadingLevel(paragraph: Paragraph): HeadingLevel | undefined {
  return paragraph.options?.heading;
}

export default function createDocxLoader(): ILoader<Buffer, Record<string, any>> {
  return createLoader({
    async pull(locale, input: Buffer) {
      const result: Record<string, any> = {};
      
      // In a real implementation, we'd parse the input buffer here
      const doc = new Document({
        sections: [{
          properties: {},
          children: []
        }]
      });

      const processTextRun = (
        textRun: TextRun, 
        path: DocElementPath
      ) => {
        const text = textRun.text || '';
        const normalizedText = normalizeTextContent(text);
        if (normalizedText) {
          result[createPathString(path)] = normalizedText;
        }
      };

      const processImageCaption = (
        paragraph: Paragraph,
        path: DocElementPath
      ) => {
        // Check if paragraph is an image caption
        // This would depend on your document structure/styling
        const isCaption = paragraph.options?.style === 'Caption';
        if (isCaption) {
          const text = paragraph.children
            .filter((child: any) => child instanceof TextRun)
            .map((child: any) => (child as TextRun).text)
            .join('');
          
          const normalizedText = normalizeTextContent(text);
          if (normalizedText) {
            result[createPathString({
              type: 'image-caption',
              indexes: path.indexes
            })] = normalizedText;
          }
        }
      };

      const processParagraph = (
        paragraph: Paragraph,
        baseIndexes: number[]
      ) => {
        const headingLevel = getHeadingLevel(paragraph);
        const path: DocElementPath = {
          type: headingLevel ? 'heading' : 'paragraph',
          indexes: [...baseIndexes]
        };

        // Process text content
        paragraph.children.forEach((child: any, idx: any) => {
          if (child instanceof TextRun) {
            processTextRun(child, {
              ...path,
              indexes: [...path.indexes, idx]
            });
          }
        });

        // Check for image captions
        processImageCaption(paragraph, path);
      };

      const processTableCell = (
        cell: TableCell,
        rowIndex: number,
        cellIndex: number,
        tableIndex: number
      ) => {
        cell.children.forEach((child: any, childIndex: any) => {
          if (child instanceof Paragraph) {
            const path: DocElementPath = {
              type: 'table-cell',
              indexes: [tableIndex, rowIndex, cellIndex, childIndex]
            };
            
            child.children.forEach((run: any, runIndex: any) => {
              if (run instanceof TextRun) {
                processTextRun(run, {
                  ...path,
                  indexes: [...path.indexes, runIndex]
                });
              }
            });
          }
        });
      };

      const processTable = (table: Table, tableIndex: number) => {
        table.rows.forEach((row: any, rowIndex: any) => {
          row.cells.forEach((cell: any, cellIndex: any) => {
            processTableCell(cell, rowIndex, cellIndex, tableIndex);
          });
        });
      };

      // Process all elements in the document
      doc.sections.forEach((section: any, sectionIndex: any) => {
        section.children.forEach((child: any, childIndex: any) => {
          if (child instanceof Paragraph) {
            processParagraph(child, [sectionIndex, childIndex]);
          } else if (child instanceof Table) {
            processTable(child, childIndex);
          }
        });
      });

      return result;
    },

    async push(locale, data: Record<string, any>, originalInput?: Buffer) {
      const doc = new Document({
        sections: [{
          properties: {},
          children: []
        }]
      });

      // Group content by type and structure
      const contentGroups: {
        paragraphs: Record<string, string[]>,
        headings: Record<string, string[]>,
        tableCells: Record<string, string[]>,
        imageCaptions: Record<string, string>
      } = {
        paragraphs: {},
        headings: {},
        tableCells: {},
        imageCaptions: {}
      };

      // Sort and group paths
      Object.entries(data).sort(([a], [b]) => {
        const aPath = parsePathString(a);
        const bPath = parsePathString(b);
        return aPath.indexes[0] - bPath.indexes[0];
      }).forEach(([path, value]) => {
        const parsedPath = parsePathString(path);
        const baseKey = parsedPath.indexes.slice(0, -1).join('/');

        switch (parsedPath.type) {
          case 'paragraph':
            if (!contentGroups.paragraphs[baseKey]) {
              contentGroups.paragraphs[baseKey] = [];
            }
            contentGroups.paragraphs[baseKey].push(value);
            break;
          case 'heading':
            if (!contentGroups.headings[baseKey]) {
              contentGroups.headings[baseKey] = [];
            }
            contentGroups.headings[baseKey].push(value);
            break;
          case 'table-cell':
            if (!contentGroups.tableCells[baseKey]) {
              contentGroups.tableCells[baseKey] = [];
            }
            contentGroups.tableCells[baseKey].push(value);
            break;
          case 'image-caption':
            contentGroups.imageCaptions[baseKey] = value;
            break;
        }
      });

      // Create paragraphs
      const createParagraph = (
        texts: string[], 
        options?: IParagraphOptions
      ): Paragraph => {
        return new Paragraph({
          ...options,
          children: texts.map(text => new TextRun({ text }))
        });
      };

      // Create table cells
      const createTableCell = (texts: string[]): TableCell => {
        return new TableCell({
          children: [createParagraph(texts)]
        });
      };

      // Build document structure
      const children: (Paragraph | Table)[] = [];

      // Add paragraphs and headings
      Object.entries(contentGroups.paragraphs).forEach(([key, texts]) => {
        children.push(createParagraph(texts));
      });

      Object.entries(contentGroups.headings).forEach(([key, texts]) => {
        const level = parseInt(key.split('/')[0]) + 1;
        children.push(createParagraph(texts, { 
          heading: HeadingLevel[`HEADING_${level}` as keyof typeof HeadingLevel] 
        }));
      });

      // Add tables
      Object.entries(contentGroups.tableCells).forEach(([key, texts]) => {
        const [tableIndex, rowIndex, cellIndex] = key.split('/').map(i => parseInt(i));
        
        if (!children[tableIndex]) {
          children[tableIndex] = new Table({
            rows: [new TableRow({
              children: [createTableCell(texts)]
            })]
          });
        } else {
          const table = children[tableIndex] as Table;
          if (!table.rows[rowIndex]) {
            table.rows[rowIndex] = new TableRow({
              children: [createTableCell(texts)]
            });
          } else {
            table.rows[rowIndex].cells[cellIndex] = createTableCell(texts);
          }
        }
      });

      // Add image captions
      Object.entries(contentGroups.imageCaptions).forEach(([key, text]) => {
        children.push(createParagraph([text], { style: 'Caption' }));
      });

      // Set document content
      doc.sections[0].children = children;

      // Convert to buffer
      const buffer = await Packer.toBuffer(doc);
      return buffer;
    }
  });
}
