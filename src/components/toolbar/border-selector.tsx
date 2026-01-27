
import { Button } from "@/components/ui/button";
import { ButtonGroup } from "@/components/ui/button-group";
import {
    DropdownMenu,
    DropdownMenuContent,
    DropdownMenuGroup,
    DropdownMenuItem,
    DropdownMenuLabel,
    DropdownMenuPortal,
    DropdownMenuSeparator,
    DropdownMenuSub,
    DropdownMenuSubContent,
    DropdownMenuSubTrigger,
    DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";
import { CellDefaultColors, CellPresetColors } from "@/constants/sheet-styles";
import { cn } from "@/lib/utils";
import {
    ChevronDown,
    CircleDashed,
    PencilLine
} from "lucide-react";
import React, { useState } from "react";


type BorderOptions = "bottom" | "top" | "left" | "right" | "clear" | "all" | "outer";

type BorderStyle = {
    lineColor: string;
    borderOption: BorderOptions;
}

type BorderSelectorOption = {
    label: string;
    value: BorderOptions;
    icon: any;
}

export function BorderSelector({ onApplyBorder }: { onApplyBorder: (borderStyle: BorderStyle) => void }) {
    const [lastUsedBorderStyle, setLastUserBorderStyle] = useState<BorderStyle>({
        lineColor: "#000000",
        borderOption: "bottom"
    });

    const optionGroups: BorderSelectorOption[][] = [
        [
            {
                label: "Bottom border",
                value: "bottom",
                icon: CircleDashed,
            },
            {
                label: "Top border",
                value: "top",
                icon: CircleDashed,
            },
            {
                label: "Left border",
                value: "left",
                icon: CircleDashed,
            },
            {
                label: "Right border",
                value: "right",
                icon: CircleDashed,
            }
        ],
        [
            {
                label: "Clear borders",
                value: "clear",
                icon: CircleDashed,
            },
            {
                label: "All borders",
                value: "all",
                icon: CircleDashed,
            },
            {
                label: "Outer borders",
                value: "outer",
                icon: CircleDashed,
            }
        ]
    ]

    const LastUsedBorderIcon = optionGroups.flat().find((option) => option.value === lastUsedBorderStyle.borderOption)?.icon

    return (
        <ButtonGroup>
            <Button
                variant="outline"
                size="sm"
                onClick={() => {
                    onApplyBorder?.(lastUsedBorderStyle);
                }}
            >
                <LastUsedBorderIcon className="size-[18px]" />
                <div className="w-1 h-4 rounded-md" style={{ backgroundColor: lastUsedBorderStyle.lineColor }} />
            </Button>
            <DropdownMenu>
                <DropdownMenuTrigger asChild>
                    <Button variant="outline" size="sm">
                        <ChevronDown />
                    </Button>
                </DropdownMenuTrigger>
                <DropdownMenuContent className="w-44" align="start" side="bottom">
                    <DropdownMenuGroup>
                        <DropdownMenuLabel>Borders</DropdownMenuLabel>
                        <DropdownMenuSeparator />
                        {
                            optionGroups.map((optGroup, idx) => {
                                const hasSeparator = idx === optionGroups.length - 1 ? false : true;
                                return (
                                    <React.Fragment key={idx}>
                                        <DropdownMenuGroup>
                                            {optGroup.map((option) => {
                                                const Icon = option.icon;
                                                return (
                                                    <DropdownMenuItem key={option.value} onClick={() => {
                                                        const newStyle = {
                                                            ...lastUsedBorderStyle,
                                                            borderOption: option.value
                                                        };
                                                        setLastUserBorderStyle(newStyle);
                                                        onApplyBorder?.(newStyle);
                                                    }}>
                                                        <Icon className="size-[18px]" />
                                                        {option.label}
                                                    </DropdownMenuItem>
                                                )
                                            })}
                                        </DropdownMenuGroup>
                                        {hasSeparator && <DropdownMenuSeparator />}
                                    </React.Fragment>
                                )
                            })
                        }
                    </DropdownMenuGroup>
                    <DropdownMenuSeparator />
                    <DropdownMenuGroup>
                        <DropdownMenuSub>
                            <DropdownMenuSubTrigger>
                                <PencilLine />
                                Line color
                            </DropdownMenuSubTrigger>
                            <DropdownMenuPortal>
                                <DropdownMenuSubContent>
                                    <PresetColorSelector
                                        className="p-2"
                                        color={lastUsedBorderStyle.lineColor}
                                        setColor={(val) => {
                                            setLastUserBorderStyle({
                                                ...lastUsedBorderStyle,
                                                lineColor: val
                                            })
                                        }}
                                    />
                                </DropdownMenuSubContent>
                            </DropdownMenuPortal>
                        </DropdownMenuSub>
                    </DropdownMenuGroup>
                </DropdownMenuContent>
            </DropdownMenu>
        </ButtonGroup>

    );
}



interface GenericColorSelectorProps {
    className?: string;
    color: string;
    setColor: (val: string) => void;
}

function PresetColorSelector({
    className,
    color,
    setColor,
}: GenericColorSelectorProps) {
    function ColorButton({ className, buttonColor }: { className?: string; buttonColor: string }) {
        return (
            <button
                className={cn(
                    "size-6 border hover:border-2 hover:border-yellow-400",
                    className,
                    color == buttonColor ? "border-2 border-orange-400" : ""
                )}
                style={{ backgroundColor: buttonColor }}
                onClick={() => setColor(buttonColor)}
            />
        );
    }

    return (
        <div className={className}>
            <p className="mb-2 text-sm font-semibold">Preset colors</p>
            <div className="flex flex-row gap-1">
                {CellPresetColors.map((colorRow, idx1) => {
                    return (
                        <div className="flex flex-col gap-0" key={idx1}>
                            {colorRow.map((color, idx2) => {
                                return (
                                    <ColorButton
                                        key={idx2}
                                        className={idx2 === 0 ? "mb-1" : ""}
                                        buttonColor={color}
                                    />
                                );
                            })}
                        </div>
                    );
                })}
            </div>
            <div className="border-t w-full my-2" />
            <p className="mb-2 text-sm font-semibold">Default colors</p>
            <div className="flex flex-row gap-1">
                {CellDefaultColors.map((color, idx) => {
                    return <ColorButton key={idx} buttonColor={color} />;
                })}
            </div>

        </div>
    );
}