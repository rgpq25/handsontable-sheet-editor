import { Workbook } from "exceljs";
import type { Cell, CellValue, Fill } from "exceljs";
import { EXCEL_THEME_COLORS, EXCEL_INDEXED_COLORS, applyTint } from "../constants/sheet-style";

export interface ExcelCellMeta {
    row: number;
    col: number;
    key: string;
    value: any;
}

export interface ExcelParsedData {
    data: any[][];
    mergeCells: { row: number; col: number; rowspan: number; colspan: number }[];
    colWidths: number[];
    rowHeights: (number | undefined)[];
    cellMeta: ExcelCellMeta[];
    customBorders: any[];
}

function excelColorToCss(
    color: { argb?: string; theme?: number; indexed?: number; tint?: number } | null | undefined
): string | undefined {
    if (!color) return undefined;

    // Handle ARGB colors
    if (color.argb) {
        const argb = color.argb;
        if (argb.length === 8) {
            const a = argb.substring(0, 2);
            const r = argb.substring(2, 4);
            const g = argb.substring(4, 6);
            const b = argb.substring(6, 8);
            if (a.toLowerCase() === "ff") {
                return `#${r}${g}${b}`;
            }
            return `#${r}${g}${b}${a}`;
        }
    }

    // Handle indexed colors
    if (color.indexed !== undefined) {
        const indexedColor = EXCEL_INDEXED_COLORS[color.indexed];
        if (indexedColor) {
            return indexedColor;
        }
    }

    // Handle theme colors
    if (color.theme !== undefined) {
        const baseColor = EXCEL_THEME_COLORS[color.theme];
        if (baseColor) {
            // Apply tint if present
            if (color.tint !== undefined && color.tint !== 0) {
                return applyTint(baseColor, color.tint);
            }
            return baseColor;
        }
    }

    return undefined;
}

function normalizeExcelValue(cell: Cell): CellValue {
    if (!cell || cell.value === undefined || cell.value === null) return null;

    const value = cell.value;

    if (typeof value === "object" && value !== null) {
        if ("result" in value) return value.result;
        if ("richText" in value) return value.richText.map((rt: any) => rt.text).join("");
        if (value instanceof Date) return value;
        if ("text" in value && "hyperlink" in value) return value.text;
    }

    return value;
}

function parseA1Range(ref: string) {
    const m = /([A-Z]+)(\d+):([A-Z]+)(\d+)/.exec(ref.toUpperCase());
    if (!m) return null;
    const lettersToIndex = (letters: string) => {
        let n = 0;
        for (let i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64);
        return n - 1; // zero-based
    };
    const c1 = lettersToIndex(m[1]);
    const r1 = parseInt(m[2], 10) - 1;
    const c2 = lettersToIndex(m[3]);
    const r2 = parseInt(m[4], 10) - 1;
    return { r1, c1, r2, c2 };
}

export async function readExcelFile(file: File): Promise<ExcelParsedData> {
    const arrBuf = await file.arrayBuffer();
    const wb = new Workbook();
    await wb.xlsx.load(arrBuf);
    const ws = wb.worksheets[0]; // $$$ TODO: Use important sheet

    const mergeCells: { row: number; col: number; rowspan: number; colspan: number }[] = [];
    const excelMerges: string[] = ws.model.merges;
    for (const ref of excelMerges) {
        const p = parseA1Range(ref);
        if (!p) continue;
        const rowspan = p.r2 - p.r1 + 1;
        const colspan = p.c2 - p.c1 + 1;
        mergeCells.push({ row: p.r1, col: p.c1, rowspan, colspan });
    }

    const MIN_ROWS = 40;
    const MIN_COLS = 26;

    const tightRow = ws.actualRowCount || ws.rowCount || 0;
    let tightCol = 0;
    for (let r = 1; r <= tightRow; r++) {
        const row = ws.getRow(r);
        tightCol = Math.max(tightCol, row.actualCellCount || 0);
    }
    const maxRow = Math.max(MIN_ROWS, tightRow);
    const maxCol = Math.max(MIN_COLS, tightCol);

    const data: any[][] = Array.from({ length: maxRow }, () => Array(maxCol).fill(null));

    const cellMeta: ExcelCellMeta[] = [];
    const customBorders: any[] = [];

    for (let r = 1; r <= maxRow; r++) {
        for (let c = 1; c <= maxCol; c++) {
            const xl = ws.getCell(r, c);

            data[r - 1][c - 1] = normalizeExcelValue(xl);

            // Font
            const f = xl.font;
            if (r === 1 && c === 1) console.log("Font: ", xl.font);
            if (f) {
                if (f.name)
                    cellMeta.push({ row: r - 1, col: c - 1, key: "fontFamily", value: f.name });
                if (f.size)
                    cellMeta.push({ row: r - 1, col: c - 1, key: "fontSize", value: f.size });
                if (f.bold) cellMeta.push({ row: r - 1, col: c - 1, key: "isBold", value: true });
                if (f.italic)
                    cellMeta.push({ row: r - 1, col: c - 1, key: "isItalic", value: true });
                if (f.underline)
                    cellMeta.push({ row: r - 1, col: c - 1, key: "isUnderline", value: true });

                const color = excelColorToCss(f.color);
                if (color)
                    cellMeta.push({ row: r - 1, col: c - 1, key: "textColor", value: color });
            }

            // Fill
            const fill = xl.fill as Fill;
            if (fill && fill.type === "pattern" && fill.pattern === "solid") {
                const bg =
                    excelColorToCss(fill.fgColor as any) || excelColorToCss(fill.bgColor as any);
                if (bg)
                    cellMeta.push({ row: r - 1, col: c - 1, key: "backgroundColor", value: bg });
            }

            // Alignment
            const a = xl.alignment;
            if (a) {
                if (a.horizontal)
                    cellMeta.push({ row: r - 1, col: c - 1, key: "hAlign", value: a.horizontal });
                if (a.vertical)
                    cellMeta.push({ row: r - 1, col: c - 1, key: "vAlign", value: a.vertical });
            }

            // Borders
            const b = xl.border;
            if (b) {
                const hotBorder: any = {
                    range: {
                        from: { row: r - 1, col: c - 1 },
                        to: { row: r - 1, col: c - 1 },
                    },
                };
                let hasBorder = false;

                if (b.top && b.top.style) {
                    hotBorder.top = { color: excelColorToCss(b.top.color) || "#000000", width: 1 };
                    hasBorder = true;
                }
                if (b.bottom && b.bottom.style) {
                    hotBorder.bottom = {
                        color: excelColorToCss(b.bottom.color) || "#000000",
                        width: 1,
                    };
                    hasBorder = true;
                }
                if (b.left && b.left.style) {
                    hotBorder.start = {
                        color: excelColorToCss(b.left.color) || "#000000",
                        width: 1,
                    };
                    hasBorder = true;
                }
                if (b.right && b.right.style) {
                    hotBorder.end = {
                        color: excelColorToCss(b.right.color) || "#000000",
                        width: 1,
                    };
                    hasBorder = true;
                }

                if (hasBorder) customBorders.push(hotBorder);
            }
        }
    }

    const colWidths: number[] = [];
    for (let c = 1; c <= maxCol; c++) {
        const colWidth = ws.getColumn(c).width;
        colWidths.push(colWidth ? Math.round(colWidth * 6.65) : 150);
    }

    const rowHeights: (number | undefined)[] = [];
    for (let r = 1; r <= maxRow; r++) {
        const hPt = ws.getRow(r).height;
        rowHeights.push(hPt ? Math.round(hPt * (96 / 72)) : undefined);
    }

    return {
        data,
        mergeCells,
        colWidths,
        rowHeights,
        cellMeta,
        customBorders,
    };
}
