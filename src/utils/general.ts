const numberType = null as unknown as number;
const booleanType = null as unknown as boolean;

export const CellNumberFormats = [
    {
        key: "decimals",
        type: numberType,
        label: "2 decimals",
        defaultValue: 0 as number,
    },
    {
        key: "commas",
        type: booleanType,
        label: "Commas (23,000)",
        defaultValue: false as boolean,
    },
    {
        key: "parens",
        type: booleanType,
        label: "Negatives as (123)",
        defaultValue: false as boolean,
    },
    {
        key: "percent",
        type: booleanType,
        label: "Percentage (%)",
        defaultValue: false as boolean,
    },
    {
        key: "zeroDash",
        type: booleanType,
        label: "Zero as -",
        defaultValue: false as boolean,
    },
] as const;

export type CellNumberFormatFlags = {
    [F in (typeof CellNumberFormats)[number] as F["key"]]: F["type"];
};

export const defaultCellNumberFormatFlags: CellNumberFormatFlags = Object.fromEntries(
    CellNumberFormats.map((f) => [f.key, f.defaultValue])
) as CellNumberFormatFlags;

export function formatWithNumFmt(value: unknown, fmt?: string): string | undefined {
    if (fmt == null || fmt === "" || fmt === "General") return undefined;
    if (value == null || value === "") return undefined;

    let num: number;
    let pattern = fmt;

    // Accept both number and numeric string
    if (typeof value === "number") {
        num = value;
    } else if (typeof value === "string") {
        // Try to parse "1234.56" or "1,234.56" as number
        const cleaned = value.replace(/\s/g, "").replace(/,/g, "");
        const parsed = Number(cleaned);
        if (!Number.isFinite(parsed)) {
            return undefined; // treat as plain text if it's not numeric
        }
        num = parsed;
    } else {
        return undefined;
    }

    // Split format sections: positive;negative;zero;text
    const sections = pattern.split(";");
    if (sections.length > 1) {
        if (num < 0 && sections[1]) {
            pattern = sections[1];
            num = Math.abs(num);
        } else if (num === 0 && sections[2]) {
            pattern = sections[2];
        } else {
            pattern = sections[0];
        }
    } else {
        pattern = sections[0];
    }

    // Percentage
    const isPercent = pattern.includes("%");
    if (isPercent) {
        num = num * 100;
    }

    // Decimal places from pattern
    const decimalPart = pattern.split(".")[1];
    const decimals = decimalPart ? decimalPart.replace(/[^0#]/g, "").length : 0;

    const useGrouping = pattern.includes(",");

    const formatted = num.toLocaleString(undefined, {
        useGrouping,
        minimumFractionDigits: decimals,
        maximumFractionDigits: decimals,
    });

    const usesParentheses = pattern.includes("(") && pattern.includes(")");
    const isNegativeOriginal =
        typeof value === "number" ? value < 0 : value.toString().trim().startsWith("-");

    let result = formatted;
    if (usesParentheses && isNegativeOriginal) {
        result = `(${formatted})`;
    }

    return isPercent ? result + "%" : result;
}

export function numFmtToFlags(fmt?: string): CellNumberFormatFlags {
    const flags: CellNumberFormatFlags = { ...defaultCellNumberFormatFlags };

    if (!fmt) return flags;

    const DECIMALS_MIN = 0;
    const DECIMALS_MAX = 6;

    const patternStr = String(fmt).trim();

    // positive ; negative ; zero ; (optional text section)
    const parts = patternStr.split(";");
    const positive = (parts[0] ?? "").trim();
    const negative = (parts[1] ?? "").trim();
    const zero = (parts[2] ?? "").trim();

    // parens negative: "(0.00)"
    if (negative.match(/^\(.*\)$/)) {
        flags.parens = true;
    }

    if (zero) {
        // direct simple forms: "-" or "\"-\""
        if (zero === "-" || zero === '"-"') {
            flags.zeroDash = true;
        } else if (zero.includes('"-"')) {
            // built-in Excel formats often have something like _(* "-"_)
            flags.zeroDash = true;
        } else {
            // normalize: strip quotes, underscores, spaces, asterisks
            const normalizedZero = zero
                .replace(/"/g, "")
                .replace(/_/g, "")
                .replace(/\*/g, "")
                .replace(/\s/g, "");

            // examples to catch:
            //   "-"          -> "-"
            //   _(*-_)       -> "(*-)" -> after strip -> "(-)" (ignore)
            //   _(* "-"_)    -> "_(*"-"_)" -> after strip -> "(*-)" (still has -)
            if (normalizedZero === "-" || normalizedZero === "-;") {
                flags.zeroDash = true;
            }
        }
    }

    // From here, parse only the positive section (same as before)
    let pattern = positive;

    // percent
    if (pattern.endsWith("%")) {
        flags.percent = true;
        pattern = pattern.slice(0, -1).trim();
    }

    // decimals: count digits after '.'
    const dotIndex = pattern.indexOf(".");
    if (dotIndex >= 0) {
        const decimalsPart = pattern.slice(dotIndex + 1);
        const digitRun = decimalsPart.match(/^\d+/)?.[0] ?? "";
        const decimalsCount = digitRun.length;

        flags.decimals = Math.min(DECIMALS_MAX, Math.max(DECIMALS_MIN, decimalsCount));
    } else {
        flags.decimals = 0;
    }

    // commas
    if (pattern.includes(",")) {
        flags.commas = true;
    }

    return flags;
}
