import { Workbook } from "exceljs";
import type Handsontable from "handsontable";
import type { ComputedBorder } from "handsontable/plugins/customBorders";

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

    return undefined;
}

export async function exportToExcel(hot: Handsontable, filename: string = "export.xlsx") {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet("Sheet 1");

    const rowCount = hot.countRows();
    const colCount = hot.countCols();

    // Set Column Widths
    for (let c = 0; c < colCount; c++) {
        const width = hot.getColWidth(c);
        if (width) {
            worksheet.getColumn(c + 1).width = width / 6.65;
        }
    }

    // Set Row Heights and Data/Styles
    for (let r = 0; r < rowCount; r++) {
        const height = hot.getRowHeight(r);
        if (height) {
            worksheet.getRow(r + 1).height = height * (72 / 96);
        }

        for (let c = 0; c < colCount; c++) {
            const cell = worksheet.getCell(r + 1, c + 1);
            const value = hot.getDataAtCell(r, c);
            cell.value = value;

            const meta = hot.getCellMeta(r, c);
            const {
                fontFamily,
                fontSize,
                isBold,
                isItalic,
                isUnderline,
                backgroundColor,
                textColor,
                hAlign,
                vAlign,
                format,
            } = meta;

            // Font
            if (isBold || isItalic || isUnderline || textColor) {
                cell.font = {
                    name: fontFamily,
                    size: fontSize,
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

            // Number Format
            if (format) {
                cell.numFmt = format;
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
    if (!mergeCellsPlugin.isEnabled()) {
        alert("Error, merge cells plugin is not enabled!");
        return;
    }

    // @ts-expect-error $$$ TODO: For some reason, the type of mergedCellsCollection is not correct
    const mergedCells = mergeCellsPlugin.mergedCellsCollection.mergedCells;
    mergedCells.forEach((merge: any) => {
        worksheet.mergeCells(
            merge.row + 1,
            merge.col + 1,
            merge.row + merge.rowspan,
            merge.col + merge.colspan
        );
    });

    // Borders (Custom Borders Plugin)
    const customBordersPlugin = hot.getPlugin("customBorders");
    if (!customBordersPlugin.isEnabled()) {
        alert("Error, custom borders plugin is not enabled!");
        return;
    }
    // @ts-expect-error savedBorders is not in the type definitions
    const savedBorders: ComputedBorder[] | undefined = customBordersPlugin.savedBorders;
    if (!savedBorders || !Array.isArray(savedBorders)) {
        alert("Error, couldnt access the savedBorders property!");
        return;
    }

    savedBorders.forEach((borderConfig: ComputedBorder) => {
        const { row, col, top, bottom, start, end } = borderConfig;

        if (row === 2 && col == 1) {
            console.log("Were on the top 'Total' cell. borderConfig =", borderConfig);
        }
        if (row === 3 && col == 1) {
            console.log("Were on the bottom 'Debit' cell. borderConfig =", borderConfig);
        }

        const cell = worksheet.getCell(row + 1, col + 1);
        const border: any = {};

        if (top && !top.hide) border.top = { style: "thin", color: cssColorToExcel(top.color) };
        if (bottom && !bottom.hide)
            border.bottom = { style: "thin", color: cssColorToExcel(bottom.color) };
        if (start && !start.hide)
            border.left = { style: "thin", color: cssColorToExcel(start.color) };
        if (end && !end.hide) border.right = { style: "thin", color: cssColorToExcel(end.color) };

        if (Object.keys(border).length > 0) {
            cell.border = border;
        }
    });

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
