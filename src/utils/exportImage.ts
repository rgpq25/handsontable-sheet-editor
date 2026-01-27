import html2canvas from "html2canvas-pro";
import Handsontable from "handsontable";

type CellStyle = {
    isBold?: boolean;
    isItalic?: boolean;
    isUnderline?: boolean;
    backgroundColor?: string;
    textColor?: string;
    hAlign?: "left" | "center" | "right";
    vAlign?: "top" | "middle" | "bottom";
    fontFamily?: string;
    fontSize?: number;
};

export async function exportToImage(hot: any, filename: string = "export.png") {
    const rows = hot.countRows();
    const cols = hot.countCols();
    if (!rows || !cols) return;

    // Capture current data
    const data = hot.getData();

    // Capture merges
    const mergePlugin = hot.getPlugin("mergeCells");
    const mergedCells = (mergePlugin?.mergedCellsCollection?.mergedCells || []).map((m: any) => ({
        row: m.row,
        col: m.col,
        rowspan: m.rowspan,
        colspan: m.colspan,
    }));

    // Capture colWidths / rowHeights
    const widthsPx = Array.from({ length: cols }, (_, c) => hot.getColWidth(c) || 100);
    const heightsPx = Array.from({ length: rows }, (_, r) => hot.getRowHeight(r) || 24);

    // Compute bounding box
    let minRow = rows;
    let maxRow = -1;
    let minCol = cols;
    let maxCol = -1;

    const markCell = (r: number, c: number) => {
        if (r < 0 || c < 0 || r >= rows || c >= cols) return;
        if (r < minRow) minRow = r;
        if (r > maxRow) maxRow = r;
        if (c < minCol) minCol = c;
        if (c > maxCol) maxCol = c;
    };

    for (let r = 0; r < rows; r++) {
        for (let c = 0; c < cols; c++) {
            const v = data[r][c];
            if (v !== null && v !== undefined && v !== "") {
                markCell(r, c);
            }
        }
    }

    for (const m of mergedCells) {
        for (let r = m.row; r < m.row + m.rowspan; r++) {
            for (let c = m.col; c < m.col + m.colspan; c++) {
                markCell(r, c);
            }
        }
    }

    if (maxRow === -1 || maxCol === -1) {
        minRow = 0;
        maxRow = 0;
        minCol = 0;
        maxCol = 0;
    }

    const effRows = maxRow - minRow + 1;
    const effCols = maxCol - minCol + 1;

    // Capture custom borders
    const customBordersPlugin = hot.getPlugin("customBorders");
    const allBorders =
        customBordersPlugin && customBordersPlugin.isEnabled()
            ? customBordersPlugin.getBorders([[0, 0, rows - 1, cols - 1]])
            : [];

    const trimmedCustomBorders = allBorders
        .filter(
            (b: any) => b.row >= minRow && b.row <= maxRow && b.col >= minCol && b.col <= maxCol
        )
        .map((b: any) => ({
            range: {
                from: { row: b.row - minRow, col: b.col - minCol },
                to: { row: b.row - minRow, col: b.col - minCol },
            },
            top: b.top,
            bottom: b.bottom,
            left: b.left,
            right: b.right,
        }));

    const trimmedData = data
        .slice(minRow, maxRow + 1)
        .map((row: any) => row.slice(minCol, maxCol + 1));
    const trimmedWidthsPx = widthsPx.slice(minCol, maxCol + 1);
    const trimmedHeightsPx = heightsPx.slice(minRow, maxRow + 1);

    // Capture styles for the trimmed area
    const trimmedStyles: (CellStyle | undefined)[][] = Array.from({ length: effRows }, (_, r) =>
        Array.from({ length: effCols }, (_, c) => {
            const meta = hot.getCellMeta(minRow + r, minCol + c);
            return {
                isBold: meta.isBold,
                isItalic: meta.isItalic,
                isUnderline: meta.isUnderline,
                backgroundColor: meta.backgroundColor,
                textColor: meta.textColor,
                hAlign: meta.hAlign,
                vAlign: meta.vAlign,
                fontFamily: meta.fontFamily,
                fontSize: meta.fontSize,
            };
        })
    );

    const trimmedMerges = mergedCells
        .map((m: any) => {
            const r1 = m.row;
            const r2 = m.row + m.rowspan - 1;
            const c1 = m.col;
            const c2 = m.col + m.colspan - 1;

            if (r2 < minRow || r1 > maxRow || c2 < minCol || c1 > maxCol) return null;

            const nr1 = Math.max(r1, minRow);
            const nc1 = Math.max(c1, minCol);
            const nr2 = Math.min(r2, maxRow);
            const nc2 = Math.min(c2, maxCol);

            return {
                row: nr1 - minRow,
                col: nc1 - minCol,
                rowspan: nr2 - nr1 + 1,
                colspan: nc2 - nc1 + 1,
            };
        })
        .filter((m: any) => m !== null);

    const container = document.createElement("div");
    container.style.position = "fixed";
    container.style.left = "-10000px";
    container.style.top = "0";
    container.style.background = "#ffffff";
    document.body.appendChild(container);

    const contentWidth = trimmedWidthsPx.reduce((acc, w) => acc + w, 0);

    const mirror = new Handsontable(container, {
        data: trimmedData,
        rowHeaders: false,
        colHeaders: false,
        licenseKey: "non-commercial-and-evaluation",
        mergeCells: trimmedMerges,
        customBorders: trimmedCustomBorders,
        colWidths: trimmedWidthsPx,
        rowHeights: trimmedHeightsPx,
        width: contentWidth,
        stretchH: "none",
        readOnly: true,
        renderAllRows: true, // render everything in DOM for screenshot
        cells: (row, col) => ({
            renderer: (
                hotInstance: any,
                td: HTMLTableCellElement,
                r: number,
                c: number,
                prop: any,
                value: any,
                cellProperties: any
            ) => {
                // First call the base renderer
                (Handsontable.renderers.getRenderer("text") as any).apply(null, [
                    hotInstance,
                    td,
                    r,
                    c,
                    prop,
                    value,
                    cellProperties,
                ]);

                // Then apply our style logic
                const st = trimmedStyles[r]?.[c];
                if (st) {
                    if (st.isBold) td.style.fontWeight = "bold";
                    if (st.isItalic) td.style.fontStyle = "italic";
                    if (st.isUnderline) td.style.textDecoration = "underline";
                    if (st.backgroundColor) td.style.backgroundColor = st.backgroundColor;
                    if (st.textColor) td.style.color = st.textColor;
                    if (st.hAlign) td.style.textAlign = st.hAlign;
                    if (st.vAlign) td.style.verticalAlign = st.vAlign;
                }
            },
        }),
    });

    mirror.render();
    await new Promise((resolve) => requestAnimationFrame(resolve));

    // Overlay for merged cells (to fix html2canvas issues with them)
    // We can also just capture the container directly if mirror.render() is sufficient
    // but the user suggested an overlay for better fidelity.

    // For now, let's try capturing the container directly first,
    // and only if it fails we add the overlay complexity.
    // Wait, the user specifically provided the overlay logic, so let's stick to it.

    const overlayRoot = document.createElement("div");
    overlayRoot.style.position = "absolute";
    overlayRoot.style.left = "0";
    overlayRoot.style.top = "0";
    overlayRoot.style.width = "100%";
    overlayRoot.style.height = "100%";
    overlayRoot.style.pointerEvents = "none";
    container.appendChild(overlayRoot);

    for (const m of trimmedMerges) {
        const owner = mirror.getCell(m.row, m.col) as HTMLTableCellElement | null;
        if (!owner) continue;

        const bottomRight =
            (mirror.getCell(
                m.row + m.rowspan - 1,
                m.col + m.colspan - 1
            ) as HTMLTableCellElement | null) || owner;

        const ownerRect = owner.getBoundingClientRect();
        const brRect = bottomRight.getBoundingClientRect();
        const containerRect = container.getBoundingClientRect();

        const left = ownerRect.left - containerRect.left;
        const top = ownerRect.top - containerRect.top;
        const width = brRect.right - ownerRect.left;
        const height = brRect.bottom - ownerRect.top;

        const st = trimmedStyles[m.row]?.[m.col];
        const csOwner = getComputedStyle(owner);
        const csBR = getComputedStyle(bottomRight);

        const overlay = document.createElement("div");
        overlay.textContent = owner.textContent || "";
        overlay.style.position = "absolute";
        overlay.style.left = `${left}px`;
        overlay.style.top = `${top}px`;
        overlay.style.width = `${width}px`;
        overlay.style.height = `${height}px`;
        overlay.style.display = "flex";
        overlay.style.alignItems =
            st?.vAlign === "top" ? "flex-start" : st?.vAlign === "bottom" ? "flex-end" : "center";
        overlay.style.justifyContent =
            st?.hAlign === "left" ? "flex-start" : st?.hAlign === "right" ? "flex-end" : "center";
        overlay.style.textAlign =
            st?.hAlign === "left" ? "left" : st?.hAlign === "right" ? "right" : "center";

        overlay.style.backgroundColor = st?.backgroundColor || csOwner.backgroundColor || "#ffffff";
        overlay.style.color = st?.textColor || csOwner.color || "#000000";

        const fontSize = (st?.fontSize ?? 11) * 1.25;
        overlay.style.fontFamily = st?.fontFamily || csOwner.fontFamily;
        overlay.style.fontSize = `${fontSize}px`;
        overlay.style.fontWeight = st?.isBold ? "700" : csOwner.fontWeight;
        overlay.style.fontStyle = st?.isItalic ? "italic" : csOwner.fontStyle;
        overlay.style.textDecoration = st?.isUnderline ? "underline" : csOwner.textDecoration;
        overlay.style.padding = "0 4px";
        overlay.style.boxSizing = "border-box";

        overlay.style.borderTop = csOwner.borderTop;
        overlay.style.borderLeft = csOwner.borderLeft;
        overlay.style.borderRight = csBR.borderRight;
        overlay.style.borderBottom = csBR.borderBottom;

        overlayRoot.appendChild(overlay);

        owner.style.color = "transparent";
        owner.style.backgroundColor = "transparent";
    }

    try {
        const canvas = await html2canvas(container, {
            backgroundColor: "#ffffff",
            scale: 2,
            logging: false,
        });

        const link = document.createElement("a");
        link.download = filename;
        link.href = canvas.toDataURL("image/png");
        link.click();
    } catch (err) {
        console.error("Failed to export image:", err);
    } finally {
        mirror.destroy();
        document.body.removeChild(container);
    }
}
