// utils/exportImage.ts
import Handsontable from "handsontable";
import html2canvas from "html2canvas-pro";

const styleRendererClone: any = function (
    this: any,
    hotInstance: Handsontable,
    td: HTMLTableCellElement,
    row: number,
    col: number,
    prop: string | number,
    value: any,
    cellProperties: Handsontable.CellProperties
) {
    Handsontable.renderers.TextRenderer.apply(this, arguments as any);

    td.style.lineHeight = "15px";
    td.style.whiteSpace = "";
    td.style.wordBreak = "";
    td.style.padding = "0px 3px";

    if (cellProperties === undefined || !cellProperties) return;

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
    } = cellProperties as any;

    if (fontFamily) td.style.fontFamily = fontFamily;
    if (fontSize) td.style.fontSize = `${fontSize * 1.3}px`;
    if (isBold) td.style.fontWeight = "bold";
    if (isItalic) td.style.fontStyle = "italic";
    if (isUnderline) td.style.textDecoration = "underline";
    if (backgroundColor) td.style.backgroundColor = backgroundColor;
    if (textColor) td.style.color = textColor;

    if (hAlign) td.style.textAlign = hAlign;
    if (vAlign) td.style.verticalAlign = vAlign;
};

Handsontable.renderers.registerRenderer("styleRendererClone", styleRendererClone);

/**
 * Compute the used range (last row/col that contains any non-empty value),
 * then expand it to include merged cells that intersect the range.
 */
function getUsedRange(hot: Handsontable) {
    const rows = hot.countRows();
    const cols = hot.countCols();

    let lastRow = -1;
    let lastCol = -1;

    // 1) Find last row/col with data
    for (let r = 0; r < rows; r++) {
        for (let c = 0; c < cols; c++) {
            const v = hot.getDataAtCell(r, c);
            const hasValue =
                v !== null && v !== undefined && !(typeof v === "string" && v.trim() === "");

            if (hasValue) {
                if (r > lastRow) lastRow = r;
                if (c > lastCol) lastCol = c;
            }
        }
    }

    // No data at all → return minimal 1x1 (or you can bail out)
    if (lastRow < 0 || lastCol < 0) {
        return { r0: 0, c0: 0, r1: 0, c1: 0 };
    }

    // 2) Expand range to include merged cells that spill beyond lastRow/lastCol
    const mergePlugin = hot.getPlugin("mergeCells") as any;
    if (mergePlugin?.mergedCellsCollection) {
        const merges = mergePlugin.mergedCellsCollection.mergedCells || [];
        for (const m of merges) {
            const mR0 = m.row;
            const mC0 = m.col;
            const mR1 = m.row + m.rowspan - 1;
            const mC1 = m.col + m.colspan - 1;

            // If merged cell intersects current used range, expand
            const intersects = mR0 <= lastRow && mC0 <= lastCol && mR1 >= 0 && mC1 >= 0;
            if (intersects) {
                if (mR1 > lastRow) lastRow = mR1;
                if (mC1 > lastCol) lastCol = mC1;
            }
        }
    }

    return { r0: 0, c0: 0, r1: lastRow, c1: lastCol };
}

function sumSizes(count: number, getSize: (i: number) => number) {
    let total = 0;
    for (let i = 0; i < count; i++) total += getSize(i);
    return total;
}

function downloadBlob(blob: Blob, filename: string) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
}

/**
 * Export the Handsontable grid as a PNG capturing ONLY the used data area,
 * independent of viewport size, including formatting, merges, and borders (best-effort).
 */
export async function exportToImage(hot: Handsontable, filename = "sheet.png") {
    // Ensure latest render before cloning settings/data
    hot.render();

    const { r0, c0, r1, c1 } = getUsedRange(hot);
    const rowCount = r1 - r0 + 1;
    const colCount = c1 - c0 + 1;

    // Extract data for used range
    const data = hot.getData(r0, c0, r1, c1);

    // Extract row/col sizes (fallback to defaults if undefined)
    const getRowHeight = (r: number) => hot.getRowHeight(r0 + r) ?? 23;
    const getColWidth = (c: number) => hot.getColWidth(c0 + c) ?? 80;

    const bodyWidth = sumSizes(colCount, getColWidth);
    const bodyHeight = sumSizes(rowCount, getRowHeight);

    const META_KEYS = [
        "fontFamily",
        "fontSize",
        "isBold",
        "isItalic",
        "isUnderline",
        "backgroundColor",
        "textColor",
        "hAlign",
        "vAlign",
    ] as const;

    const cell: any[] = [];
    for (let r = r0; r <= r1; r++) {
        for (let c = c0; c <= c1; c++) {
            const meta = hot.getCellMeta(r, c) as any;
            const entry: any = { row: r - r0, col: c - c0 };
            let hasAny = false;

            for (const k of META_KEYS) {
                if (meta[k] !== undefined) {
                    entry[k] = meta[k];
                    hasAny = true;
                }
            }

            if (hasAny) cell.push(entry);
        }
    }

    // Best-effort merges within used range
    const mergeCells: any[] = [];
    const mergePlugin = hot.getPlugin("mergeCells") as any;
    if (mergePlugin?.mergedCellsCollection) {
        const merges = mergePlugin.mergedCellsCollection.mergedCells || [];
        for (const m of merges) {
            const mR0 = m.row;
            const mC0 = m.col;
            const mR1 = m.row + m.rowspan - 1;
            const mC1 = m.col + m.colspan - 1;

            // Include merges that intersect used range
            const intersects = mR0 <= r1 && mC0 <= c1 && mR1 >= r0 && mC1 >= c0;

            if (intersects) {
                mergeCells.push({
                    row: Math.max(mR0, r0) - r0,
                    col: Math.max(mC0, c0) - c0,
                    rowspan: m.rowspan - Math.max(0, r0 - mR0) - Math.max(0, mR1 - r1),
                    colspan: m.colspan - Math.max(0, c0 - mC0) - Math.max(0, mC1 - c1),
                });
            }
        }
    }
    const settings = hot.getSettings() as any;
    const customBorders = settings.customBorders;

    const BLEED_PX = 1;

    // Create offscreen wrapper
    const wrapper = document.createElement("div");
    wrapper.style.position = "fixed";
    wrapper.style.left = "-100000px";
    wrapper.style.top = "0";
    wrapper.style.background = "white";
    wrapper.style.zIndex = "100000";

    wrapper.style.boxSizing = "content-box";
    wrapper.style.padding = `${BLEED_PX}px`;
    wrapper.style.overflow = "visible";

    // Important: size large enough so nothing scrolls inside
    wrapper.style.width = `${bodyWidth}px`;
    wrapper.style.height = `${bodyHeight}px`;
    document.body.appendChild(wrapper);

    // Create clone container
    const cloneHost = document.createElement("div");
    wrapper.appendChild(cloneHost);

    // Build clone HOT (Core, not React)
    const clone = new Handsontable(cloneHost, {
        // themeName: "ht-theme-main", $$$ TODO Commented for now because it adds rounded corners to the table
        data,
        cell: cell,
        rowHeaders: false,
        colHeaders: false,
        renderAllRows: true,
        renderAllColumns: true,
        autoRowSize: false,
        autoColumnSize: false,
        manualRowResize: false,
        manualColumnResize: false,
        rowHeights: (index: number) => getRowHeight(index),
        colWidths: (index: number) => getColWidth(index),
        cells: () => ({ renderer: styleRendererClone }),
        mergeCells,
        customBorders,
        licenseKey: settings.licenseKey ?? "non-commercial-and-evaluation",
        outsideClickDeselects: false,
    });

    // Force layout/render
    clone.render();
    await new Promise<void>((res) => requestAnimationFrame(() => res()));

    const captureTarget = cloneHost.querySelector(".ht_master") as HTMLElement | null;
    if (!captureTarget) throw new Error("Could not find capture target");
    const rect = captureTarget.getBoundingClientRect();
    wrapper.style.width = `${Math.ceil(rect.width)}px`;
    wrapper.style.height = `${Math.ceil(rect.height)}px`;

    // Capture
    const canvas = await html2canvas(wrapper, {
        backgroundColor: "#ffffff",
        scale: Math.max(2, window.devicePixelRatio || 1),
        logging: false,
    });

    // Convert to PNG + download
    const blob: Blob | null = await new Promise((resolve) =>
        canvas.toBlob((b) => resolve(b), "image/png")
    );
    if (!blob) {
        clone.destroy();
        wrapper.remove();
        throw new Error("Failed to create PNG blob from canvas.");
    }

    downloadBlob(blob, filename);

    // Cleanup
    clone.destroy();
    wrapper.remove();
}
