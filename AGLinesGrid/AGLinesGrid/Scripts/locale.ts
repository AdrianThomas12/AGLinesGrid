type UserLanguage = {
    lcid: number;
    language_name: string;
    bcp47_locale: string; // e.g. "en-US"
    date_format: string; // e.g. "MM/dd/yyyy"
    time_format: string; // e.g. "hh:mm tt"
    thousand_separator: string;
    decimal_separator: string;
};

type UserTimezone = {
    timezonecode: number;
    userinterfacename: string;
    windows_standard_name: string;
    iana: string; // e.g. "Asia/Beirut"
};

type NumberFormatOverride = {
    decimalSeparator: string;
    groupSeparator: string;
};

let pcfNumberFormat: NumberFormatOverride | null = null;

/**
 * Pre-initialises number formatters using the PCF context's NumberFormattingInfo.
 * Call this from renderer.ts before ReactDOM.render so the correct D365 user separators
 * are used instead of falling back to navigator.language.
 */
export function initFormatters(numberFormattingInfo: {
    numberDecimalSeparator: string;
    numberGroupSeparator: string;
}) {
    const decSep = numberFormattingInfo.numberDecimalSeparator;
    const grpSep = numberFormattingInfo.numberGroupSeparator;
    if (pcfNumberFormat?.decimalSeparator !== decSep || pcfNumberFormat?.groupSeparator !== grpSep) {
        pcfNumberFormat = { decimalSeparator: decSep, groupSeparator: grpSep };
        cachedFormatters = null; // Invalidate cache so next buildFormatters() rebuilds with new separators
    }
}

function getPortalLocale() {
    const lang =
        (window as any)?.msdyn?.Portal?.Snippets?.userLanguage ||
        ((parent.window as any)?.msdyn?.Portal?.Snippets?.userLanguage as UserLanguage | undefined);
    const tz =
        (window as any)?.msdyn?.Portal?.Snippets?.userTimezone ||
        ((parent.window as any)?.msdyn?.Portal?.Snippets?.userTimezone as UserTimezone | undefined);

    const locale = lang?.bcp47_locale || navigator.language || "en-US";
    const timeZone = tz?.iana || Intl.DateTimeFormat().resolvedOptions().timeZone || "UTC";
    const datePattern = lang?.date_format || "MM/dd/yyyy";
    const timePattern = lang?.time_format || "hh:mm tt";

    return { locale, timeZone, datePattern, timePattern, lang, tz };
}

function patternToDateOptions(datePattern: string): Intl.DateTimeFormatOptions {
    // supports common tokens: dd, d, MM, M, yyyy, yy
    const twoDigit = (token: string, t2: string) =>
        datePattern.includes(token) || datePattern.includes(t2) ? ("2-digit" as const) : undefined;

    const month = datePattern.includes("MMMM")
        ? "long"
        : datePattern.includes("MMM")
        ? "short"
        : twoDigit("MM", "M")
        ? "2-digit"
        : "numeric";

    const day = datePattern.includes("dd") ? "2-digit" : "numeric";
    const year = datePattern.includes("yyyy") ? "numeric" : "2-digit";

    return { year, month: month as any, day: day as any };
}

function patternToTimeOptions(timePattern: string): Intl.DateTimeFormatOptions {
    // supports: hh, h, HH, H, mm, tt
    const hour12 = timePattern.toLowerCase().includes("tt");
    const hour =
        timePattern.includes("HH") || timePattern.includes("H")
            ? "2-digit"
            : timePattern.includes("hh") || timePattern.includes("h")
            ? "2-digit"
            : undefined;
    const minute = timePattern.includes("mm") ? "2-digit" : undefined;

    return { hour12, hour, minute } as Intl.DateTimeFormatOptions;
}

let cachedFormatters: {
    formatDate: (d: Date | string | null | undefined) => string;
    formatTime: (d: Date | string | null | undefined) => string;
    formatDateTime: (d: Date | string | null | undefined) => string;
    formatInteger: (n: any) => string;
    formatDecimal: (n: any, digits?: number) => string;
    parseDecimal: (val: string | number | null | undefined) => number;
    parseUserInput: (val: string) => number;
} = null;

export function buildFormatters() {
    if (cachedFormatters) return cachedFormatters;
    const { locale, timeZone, datePattern, timePattern, lang } = getPortalLocale();

    // Separators: PCF context (most accurate) > portal lang snippets > Intl locale detection
    const decSep =
        pcfNumberFormat?.decimalSeparator ??
        lang?.decimal_separator?.trim() ??
        new Intl.NumberFormat(locale, { minimumFractionDigits: 1 }).format(1.1)[1];
    const grpSep =
        pcfNumberFormat?.groupSeparator ??
        lang?.thousand_separator?.trim() ??
        new Intl.NumberFormat(locale, { useGrouping: true }).format(1111)[1];

    const dateFmt = new Intl.DateTimeFormat(locale, { timeZone, ...patternToDateOptions(datePattern) });
    const timeFmt = new Intl.DateTimeFormat(locale, { timeZone, ...patternToTimeOptions(timePattern) });
    const dateTimeFmt = new Intl.DateTimeFormat(locale, {
        timeZone,
        ...patternToDateOptions(datePattern),
        ...patternToTimeOptions(timePattern)
    });

    // Format a number using the resolved separators, bypassing Intl.NumberFormat locale guessing.
    function formatNum(n: number, digits: number): string {
        if (!isFinite(n)) return String(n);
        const isNeg = n < 0;
        const fixed = Math.abs(n).toFixed(digits);
        const [intPart, fracPart] = fixed.split(".");
        const intFormatted = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, grpSep);
        const result = digits > 0 ? `${intFormatted}${decSep}${fracPart}` : intFormatted;
        return isNeg ? `-${result}` : result;
    }

    cachedFormatters = {
        formatDate: (d: Date | string | null | undefined) => {
            if (!d) return "";
            const dt = d instanceof Date ? d : new Date(d);
            return isNaN(dt.getTime()) ? "" : dateFmt.format(dt);
        },
        formatTime: (d: Date | string | null | undefined) => {
            if (!d) return "";
            const dt = d instanceof Date ? d : new Date(d);
            return isNaN(dt.getTime()) ? "" : timeFmt.format(dt);
        },
        formatDateTime: (d: Date | string | null | undefined) => {
            if (!d) return "";
            const dt = d instanceof Date ? d : new Date(d);
            return isNaN(dt.getTime()) ? "" : dateTimeFmt.format(dt);
        },
        formatInteger: (n: any) =>
            n == null || n === "" || isNaN(Number(n)) ? "" : formatNum(Math.round(Number(n)), 0),
        formatDecimal: (n: any, digits = 2) =>
            n == null || n === "" || isNaN(Number(n)) ? "" : formatNum(Number(n), digits),
        parseDecimal: (val: string | number | null | undefined): number => {
            if (val == null || val === "") return 0;
            if (typeof val === "number") return isNaN(val) ? 0 : val;
            const stripped = String(val).replaceAll(grpSep, "").replace(decSep, ".");
            const n = parseFloat(stripped);
            return isNaN(n) ? 0 : n;
        },
        parseUserInput: (val: string): number => {
            if (!val || val.trim() === "") return NaN;
            // Strip spaces (\u00A0 = Finnish/French/Swedish) and apostrophes (Swiss) first
            const s = val.trim().replace(/[\s\u00A0']/g, "");
            if (!s) return NaN;

            const hasDot = s.includes(".");
            const hasComma = s.includes(",");

            // When both separators are present the last one is unambiguously the decimal
            if (hasDot && hasComma) {
                const lastDot = s.lastIndexOf(".");
                const lastComma = s.lastIndexOf(",");
                if (lastDot > lastComma) {
                    return parseFloat(s.replaceAll(",", ""));         // dot is decimal: "1,000.53"
                } else {
                    return parseFloat(s.replaceAll(".", "").replace(",", ".")); // comma is decimal: "1.000,53"
                }
            }

            // Single separator — use resolved grpSep to decide
            if (hasDot) {
                if (grpSep === ".") {
                    return parseFloat(s.replaceAll(".", "")); // dot is thousands in this locale
                } else {
                    return parseFloat(s);                     // dot is decimal in this locale
                }
            }
            if (hasComma) {
                if (grpSep === ",") {
                    return parseFloat(s.replaceAll(",", "")); // comma is thousands in this locale
                } else {
                    return parseFloat(s.replace(",", "."));   // comma is decimal in this locale
                }
            }
            return parseFloat(s);
        }
    };

    return cachedFormatters;
}