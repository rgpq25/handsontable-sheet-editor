import { Bold, Italic, Underline } from "lucide-react";
import { BorderSelector } from "./toolbar/border-selector";
import { Button } from "./ui/button";
import { ButtonGroup } from "./ui/button-group";

type BorderOptions = "bottom" | "top" | "left" | "right" | "clear" | "all" | "outer";

type BorderStyle = {
    lineColor: string;
    borderOption: BorderOptions;
}

export function Toolbar({
    onBold,
    onItalic,
    onUnderline,
    onApplyBG,
    onApplyColor,
    onApplyBorder,
    onImport,
    onExportExcel,
    onExportImage
}: {
    onBold: () => void,
    onItalic: () => void,
    onUnderline: () => void,
    onApplyBG: () => void,
    onApplyColor: () => void,
    onApplyBorder: (borderStyle: BorderStyle) => void,
    onImport: () => void,
    onExportExcel: () => void,
    onExportImage: () => void
}) {
    return (
        <div className="flex flex-row w-full h-auto border border-red-500 py-5 gap-2 px-5 items-center">
            <ButtonGroup>
                <Button variant="outline" onClick={onBold}>
                    <Bold />
                </Button>
                <Button variant="outline" onClick={onItalic}>
                    <Italic />
                </Button>
                <Button variant="outline" onClick={onUnderline}>
                    <Underline />
                </Button>
            </ButtonGroup>

            <Button variant={'outline'} onClick={onApplyBG}>Apply BG <div className="size-3 rounded-full bg-red-500" /></Button>
            <Button variant={'outline'} onClick={onApplyColor}>Apply Font <div className="size-3 rounded-full bg-green-500" /></Button>

            <BorderSelector onApplyBorder={onApplyBorder} />

            <div className="flex-1" />

            <Button variant="outline" onClick={onImport}>
                Import Excel
            </Button>
            <Button variant="outline" onClick={onExportImage}>
                Export image
            </Button>
            <Button variant="outline" onClick={onExportExcel}>
                Export excel
            </Button>
        </div>
    )
}