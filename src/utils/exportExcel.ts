import { Workbook } from "exceljs";

function cssColorToExcel(cssColor: string | undefined): { argb: string } | undefined {
    if (!cssColor) return undefined;

    // Handle hex colors (#RRGGBB or #RGB)
    if (cssColor.startsWith("#")) {
        let hex = cssColor.slice(1);
        if (hex.length === 3) {
            hex = hex
                .split("")
                .map((c) => c + c)
                .join("");
        }
        // ExcelJS expects ARGB. We'll use FF for full opacity if not provided.
        if (hex.length === 6) {
            return { argb: "FF" + hex.toUpperCase() };
        }
        if (hex.length === 8) {
            // CSS is RRGGBBAA, but ExcelJS is AARRGGBB
            const rr = hex.substring(0, 2);
            const gg = hex.substring(2, 4);
            const bb = hex.substring(4, 6);
            const aa = hex.substring(6, 8);
            return { argb: (aa + rr + gg + bb).toUpperCase() };
        }
    }

    // Basic support for some named colors or rgb (could be expanded)
    // For now, if it's not hex, we might just ignore or return a default
    return undefined;
}

export async function exportToExcel(hot: any, filename: string = "export.xlsx") {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet("Sheet 1");

    const rowCount = hot.countRows();
    const colCount = hot.countCols();

    // Set Column Widths
    for (let c = 0; c < colCount; c++) {
        const width = hot.getColWidth(c);
        if (width) {
            // Handsontable width to ExcelJS width (approximate)
            worksheet.getColumn(c + 1).width = width / 8;
        }
    }

    // Set Row Heights and Data/Styles
    for (let r = 0; r < rowCount; r++) {
        const height = hot.getRowHeight(r);
        if (height) {
            worksheet.getRow(r + 1).height = height * (72 / 96); // px to pt
        }

        for (let c = 0; c < colCount; c++) {
            const cell = worksheet.getCell(r + 1, c + 1);
            const value = hot.getDataAtCell(r, c);
            cell.value = value;

            const meta = hot.getCellMeta(r, c);
            const { isBold, isItalic, isUnderline, backgroundColor, textColor, hAlign, vAlign } =
                meta;

            // Font
            if (isBold || isItalic || isUnderline || textColor) {
                cell.font = {
                    bold: !!isBold,
                    italic: !!isItalic,
                    underline: !!isUnderline,
                    color: cssColorToExcel(textColor),
                };
            }

            // Fill
            if (backgroundColor) {
                cell.fill = {
                    type: "pattern",
                    pattern: "solid",
                    fgColor: cssColorToExcel(backgroundColor),
                };
            }

            // Alignment
            if (hAlign || vAlign) {
                cell.alignment = {
                    horizontal:
                        hAlign === "left" || hAlign === "center" || hAlign === "right"
                            ? hAlign
                            : undefined,
                    vertical:
                        vAlign === "top" || vAlign === "middle" || vAlign === "bottom"
                            ? vAlign
                            : undefined,
                    wrapText: true,
                };
            }
        }
    }

    // Merged Cells
    const mergeCellsPlugin = hot.getPlugin("mergeCells");
    if (mergeCellsPlugin && mergeCellsPlugin.isEnabled()) {
        const mergedCells = mergeCellsPlugin.mergedCellsCollection.mergedCells;
        mergedCells.forEach((merge: any) => {
            worksheet.mergeCells(
                merge.row + 1,
                merge.col + 1,
                merge.row + merge.rowspan,
                merge.col + merge.colspan
            );
        });
    }

    // Borders (Custom Borders Plugin)
    const customBordersPlugin = hot.getPlugin("customBorders");
    if (customBordersPlugin && customBordersPlugin.isEnabled()) {
        // Unfortunately, customBorders doesn't expose a simple way to get 'all' borders
        // so we iterate through cells. This can be slow for large grids.
        for (let r = 0; r < rowCount; r++) {
            for (let c = 0; c < colCount; c++) {
                const borders = customBordersPlugin.getBorders([[r, c, r, c]])[0];
                if (borders) {
                    const cell = worksheet.getCell(r + 1, c + 1);
                    const border: any = {};
                    if (borders.top)
                        border.top = { style: "thin", color: cssColorToExcel(borders.top.color) };
                    if (borders.bottom)
                        border.bottom = {
                            style: "thin",
                            color: cssColorToExcel(borders.bottom.color),
                        };
                    if (borders.left || borders.start)
                        border.left = {
                            style: "thin",
                            color: cssColorToExcel(borders.left?.color || borders.start?.color),
                        };
                    if (borders.right || borders.end)
                        border.right = {
                            style: "thin",
                            color: cssColorToExcel(borders.right?.color || borders.end?.color),
                        };

                    if (Object.keys(border).length > 0) {
                        cell.border = border;
                    }
                }
            }
        }
    }

    // Write to buffer and download
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = window.URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = filename;
    anchor.click();
    window.URL.revokeObjectURL(url);
}
