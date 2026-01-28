/**
 * Excel Color Mappings
 *
 * These constants map Excel's theme colors and indexed colors to CSS hex values.
 * Used when reading Excel files to convert Excel color specifications to web colors.
 */

/**
 * Default Office Theme Colors (Excel 2013+)
 * Maps theme index to hex color value
 */
export const EXCEL_THEME_COLORS: Record<number, string> = {
    0: "#000000", // Dark 1 (typically black)
    1: "#FFFFFF", // Light 1 (typically white)
    2: "#44546A", // Dark 2
    3: "#E7E6E6", // Light 2
    4: "#5B9BD5", // Accent 1
    5: "#ED7D31", // Accent 2
    6: "#A5A5A5", // Accent 3
    7: "#FFC000", // Accent 4
    8: "#4472C4", // Accent 5
    9: "#70AD47", // Accent 6
};

/**
 * Excel Indexed Colors (Standard 56-color palette)
 * Maps color index (1-56) to hex color value
 */
export const EXCEL_INDEXED_COLORS: Record<number, string> = {
    1: "#000000", // Black
    2: "#FFFFFF", // White
    3: "#FF0000", // Red
    4: "#00FF00", // Green
    5: "#0000FF", // Blue
    6: "#FFFF00", // Yellow
    7: "#FF00FF", // Magenta
    8: "#00FFFF", // Cyan
    9: "#800000", // Dark Red
    10: "#008000", // Dark Green
    11: "#000080", // Dark Blue
    12: "#808000", // Dark Yellow / Olive
    13: "#800080", // Dark Magenta / Purple
    14: "#008080", // Dark Cyan / Teal
    15: "#C0C0C0", // Light Gray / Silver
    16: "#808080", // Gray
    17: "#9999FF",
    18: "#993366",
    19: "#FFFFCC",
    20: "#CCFFFF",
    21: "#660066",
    22: "#FF8080",
    23: "#0066CC",
    24: "#CCCCFF",
    25: "#000080",
    26: "#FF00FF",
    27: "#FFFF00",
    28: "#00FFFF",
    29: "#800080",
    30: "#800000",
    31: "#008080",
    32: "#0000FF",
    33: "#00CCFF",
    34: "#CCFFFF",
    35: "#CCFFCC",
    36: "#FFFF99",
    37: "#99CCFF",
    38: "#FF99CC",
    39: "#CC99FF",
    40: "#FFCC99",
    41: "#3366FF",
    42: "#33CCCC",
    43: "#99CC00",
    44: "#FFCC00",
    45: "#FF9900",
    46: "#FF6600",
    47: "#666699",
    48: "#969696",
    49: "#003366",
    50: "#339966",
    51: "#003300",
    52: "#333300",
    53: "#993300",
    54: "#993366",
    55: "#333399",
    56: "#333333",
    // Special indices often used by Excel
    64: "#000000", // Default foreground (system)
};

/**
 * Apply tint to a hex color
 * Tint > 0: lighten towards white
 * Tint < 0: darken towards black
 */
export function applyTint(hexColor: string, tint: number): string {
    // Parse hex to RGB
    const hex = hexColor.replace("#", "");
    let r = parseInt(hex.substring(0, 2), 16);
    let g = parseInt(hex.substring(2, 4), 16);
    let b = parseInt(hex.substring(4, 6), 16);

    if (tint > 0) {
        // Lighten: blend towards white
        r = Math.round(r + (255 - r) * tint);
        g = Math.round(g + (255 - g) * tint);
        b = Math.round(b + (255 - b) * tint);
    } else if (tint < 0) {
        // Darken: blend towards black
        const factor = 1 + tint; // tint is negative, so this is < 1
        r = Math.round(r * factor);
        g = Math.round(g * factor);
        b = Math.round(b * factor);
    }

    // Clamp values
    r = Math.max(0, Math.min(255, r));
    g = Math.max(0, Math.min(255, g));
    b = Math.max(0, Math.min(255, b));

    return `#${r.toString(16).padStart(2, "0")}${g.toString(16).padStart(2, "0")}${b.toString(16).padStart(2, "0")}`.toUpperCase();
}
