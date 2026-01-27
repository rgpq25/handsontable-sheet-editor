import { HotTable } from '@handsontable/react-wrapper';
import Handsontable from 'handsontable';
import { useRef } from 'react';
import { Toolbar } from './components/toolbar';
import { Button } from './components/ui/button';
import { readExcelFile } from './utils/readExcel';
import { exportToExcel } from './utils/exportExcel';
import { exportToImage } from './utils/exportImage';

type BorderOptions = "bottom" | "top" | "left" | "right" | "clear" | "all" | "outer" | "inner";

type BorderStyle = {
	lineColor: string;
	borderOption: BorderOptions;
}

const styleRenderer: any = function (
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

	if (cellProperties === undefined || !cellProperties) return;

	const { isBold, isItalic, isUnderline, backgroundColor, textColor, hAlign, vAlign } = cellProperties as any;

	if (isBold) td.style.fontWeight = 'bold';
	if (isItalic) td.style.fontStyle = 'italic';
	if (isUnderline) td.style.textDecoration = 'underline';
	if (backgroundColor) td.style.backgroundColor = backgroundColor;
	if (textColor) td.style.color = textColor;

	if (hAlign) td.style.textAlign = hAlign;
	if (vAlign) td.style.verticalAlign = vAlign;
}

Handsontable.renderers.registerRenderer("styleRenderer", styleRenderer);



function App() {
	const hotTableComponent = useRef<any>(null);
	const fileInputRef = useRef<HTMLInputElement>(null);

	const handleFileUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
		const file = event.target.files?.[0];
		if (!file) return;

		const parsedData = await readExcelFile(file);
		const hot: Handsontable | null = hotTableComponent.current?.hotInstance;

		if (hot) {
			// Group cell values by row/col for 'cell' option
			const cellSettings: any[] = [];
			const cellMetaMap: Record<string, any> = {};

			parsedData.cellMeta.forEach(meta => {
				const key = `${meta.row},${meta.col}`;
				if (!cellMetaMap[key]) {
					cellMetaMap[key] = { row: meta.row, col: meta.col };
				}
				cellMetaMap[key][meta.key] = meta.value;
			});

			Object.values(cellMetaMap).forEach(val => cellSettings.push(val));

			hot.updateSettings({
				data: parsedData.data,
				mergeCells: parsedData.mergeCells,
				colWidths: parsedData.colWidths,
				rowHeights: parsedData.rowHeights,
				cell: cellSettings,
				customBorders: parsedData.customBorders
			});
		}
	};

	const handleExportExcel = async () => {
		const hot: Handsontable | null = hotTableComponent.current?.hotInstance;
		if (hot) {
			await exportToExcel(hot, "handsontable_export.xlsx");
		}
	};

	const handleExportImage = async () => {
		const hot: Handsontable | null = hotTableComponent.current?.hotInstance;
		if (hot) {
			await exportToImage(hot, "handsontable_screenshot.png");
		}
	};

	const applyStyle = (key: string, value: any, toggle: boolean = false) => {
		const hot: Handsontable | null = hotTableComponent.current?.hotInstance;
		if (!hot) return;

		const selected = hot.getSelected();
		if (!selected) return;

		selected.forEach((range: any) => {
			const [r1, c1, r2, c2] = range;
			for (let r = Math.min(r1, r2); r <= Math.max(r1, r2); r++) {
				for (let c = Math.min(c1, c2); c <= Math.max(c1, c2); c++) {
					let newValue = value;
					if (toggle) {
						const currentMeta = hot.getCellMeta(r, c);
						newValue = !currentMeta[key];
					}
					hot.setCellMeta(r, c, key, newValue);
				}
			}
		});
		hot.render();
	};

	const applyBorder = (borderStyle: BorderStyle) => {
		const hot: Handsontable | null = hotTableComponent.current?.hotInstance;
		if (!hot) return;
		const selected = hot.getSelected();
		if (!selected) return;
		const customBordersPlugin = hot.getPlugin('customBorders');
		if (!customBordersPlugin || !customBordersPlugin.isEnabled()) return;

		const { lineColor, borderOption } = borderStyle;

		hot.suspendRender();
		switch (borderOption) {
			case "bottom":
				customBordersPlugin.setBorders(selected, { bottom: { color: lineColor, width: 1 } });
				break;
			case "top":
				customBordersPlugin.setBorders(selected, { top: { color: lineColor, width: 1 } });
				break;
			case "left":
				customBordersPlugin.setBorders(selected, { start: { color: lineColor, width: 1 } });
				break;
			case "right":
				customBordersPlugin.setBorders(selected, { end: { color: lineColor, width: 1 } });
				break;
			case "clear":
				selected.forEach((range: any) => {
					const [r1, c1, r2, c2] = range;
					const top = Math.min(r1, r2);
					const bottom = Math.max(r1, r2);
					const left = Math.min(c1, c2);
					const right = Math.max(c1, c2);

					customBordersPlugin.clearBorders([range]);

					const mergeCellsPlugin = hot.getPlugin('mergeCells') as any;

					const getEffectiveRange = (r: number, c: number): [number, number, number, number][] => {
						const merged = mergeCellsPlugin.mergedCellsCollection.get(r, c);
						return merged
							? [[merged.row, merged.col, merged.row + merged.rowspan - 1, merged.col + merged.colspan - 1]]
							: [[r, c, r, c]];
					};
					if (top > 0) {
						for (let c = left; c <= right; c++) {
							customBordersPlugin.setBorders(getEffectiveRange(top - 1, c), { bottom: { hide: true } });
						}
					}
					if (bottom < hot.countRows() - 1) {
						for (let c = left; c <= right; c++) {
							customBordersPlugin.setBorders(getEffectiveRange(bottom + 1, c), { top: { hide: true } });
						}
					}
					if (left > 0) {
						for (let r = top; r <= bottom; r++) {
							customBordersPlugin.setBorders(getEffectiveRange(r, left - 1), { end: { hide: true } });
						}
					}
					if (right < hot.countCols() - 1) {
						for (let r = top; r <= bottom; r++) {
							customBordersPlugin.setBorders(getEffectiveRange(r, right + 1), { start: { hide: true } });
						}
					}
				});
				break;
			case "all":
				customBordersPlugin.setBorders(selected, { bottom: { color: lineColor, width: 1 }, top: { color: lineColor, width: 1 }, start: { color: lineColor, width: 1 }, end: { color: lineColor, width: 1 } });
				break;
			case "outer":
				selected.forEach((range: any) => {
					const [r1, c1, r2, c2] = range;
					const top = Math.min(r1, r2);
					const bottom = Math.max(r1, r2);
					const left = Math.min(c1, c2);
					const right = Math.max(c1, c2);
					customBordersPlugin.setBorders([[top, left, top, right]], { top: { color: lineColor, width: 1 } });
					customBordersPlugin.setBorders([[bottom, left, bottom, right]], { bottom: { color: lineColor, width: 1 } });
					customBordersPlugin.setBorders([[top, left, bottom, left]], { start: { color: lineColor, width: 1 } });
					customBordersPlugin.setBorders([[top, right, bottom, right]], { end: { color: lineColor, width: 1 } });
				});
				break;
		}
		hot.resumeRender();
		hot.render();
	};

	return (
		<div className='w-full h-dvh flex flex-col p-10 gap-5'>
			<Toolbar
				onBold={() => applyStyle('isBold', true, true)}
				onItalic={() => applyStyle('isItalic', true, true)}
				onUnderline={() => applyStyle('isUnderline', true, true)}
				onApplyBG={() => applyStyle('backgroundColor', '#ef4444')} // red-500
				onApplyColor={() => applyStyle('textColor', '#22c55e')} // green-500
				onApplyBorder={(borderStyle: BorderStyle) => applyBorder(borderStyle)}
				onImport={() => fileInputRef.current?.click()}
				onExportExcel={handleExportExcel}
				onExportImage={handleExportImage}
			/>
			<input
				type="file"
				ref={fileInputRef}
				onChange={handleFileUpload}
				accept=".xlsx"
				className="hidden"
			/>
			<Button onClick={() => {
				const hot: Handsontable | null = hotTableComponent.current?.hotInstance;
				if (!hot) return;
				const selected = hot.getSelected();
				if (!selected) return;
				const customBordersPlugin = hot.getPlugin('customBorders');
				if (!customBordersPlugin || !customBordersPlugin.isEnabled()) return;

				console.log("Selected: ", selected)
				console.log("Borders: ", customBordersPlugin.getBorders(selected))
			}}>
				Inspect borders
			</Button>
			<HotTable
				ref={hotTableComponent}
				themeName='ht-theme-main'
				key={"ht-theme-main"}
				className="h-full w-full"
				rowHeaders={true}
				colHeaders={true}
				height="auto"
				autoRowSize={false}
				autoColumnSize={false}
				manualRowResize={true}
				manualColumnResize={true}
				outsideClickDeselects={false}
				cells={() => ({ renderer: styleRenderer })}
				licenseKey="non-commercial-and-evaluation"
				customBorders={true}
				mergeCells={true}
				contextMenu={{
					items: {
						row_above: {},
						row_below: {},
						col_left: {},
						col_right: {},
						sp1: '---------',
						remove_row: {},
						remove_col: {},
						clear_column: {},
						sp2: '---------',
						copy_styles: {
							name: "Copy styles",
							callback: () => { console.log("Copying styles!") }
						},
						paste_styles: {
							name: "Paste styles",
							callback() { console.log("Pasting styles!") },
						},
						sp3: '---------',
						mergeCells: {},

					}
				}}
			/>
		</div>
	)
}

export default App
