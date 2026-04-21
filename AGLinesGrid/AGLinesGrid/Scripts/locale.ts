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

    const dateFmt = new Intl.DateTimeFormat(locale, { timeZone, ...patternToDateOptions(datePattern) });
    const timeFmt = new Intl.DateTimeFormat(locale, { timeZone, ...patternToTimeOptions(timePattern) });
    const dateTimeFmt = new Intl.DateTimeFormat(locale, {
        timeZone,
        ...patternToDateOptions(datePattern),
        ...patternToTimeOptions(timePattern)
    });

    // generic number formatters (optional, but nice)
    const intFmt = new Intl.NumberFormat(locale, { maximumFractionDigits: 0 });
    const decFmt = (fractionDigits = 2) =>
        new Intl.NumberFormat(locale, { minimumFractionDigits: fractionDigits, maximumFractionDigits: fractionDigits });

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
        formatInteger: (n: any) => (n == null || n === "" || isNaN(Number(n)) ? "" : intFmt.format(Number(n))),
        formatDecimal: (n: any, digits = 2) =>
            n == null || n === "" || isNaN(Number(n)) ? "" : decFmt(digits).format(Number(n)),
        parseDecimal: (val: string | number | null | undefined): number => {
            if (val == null || val === "") return 0;
            if (typeof val === "number") return isNaN(val) ? 0 : val;
            const groupingSep =
                lang?.thousand_separator?.trim() ??
                new Intl.NumberFormat(locale, { useGrouping: true }).format(1111)[1];
            const decimalSep =
                lang?.decimal_separator?.trim() ??
                new Intl.NumberFormat(locale, { minimumFractionDigits: 1 }).format(1.1)[1];
            const stripped = String(val).replaceAll(groupingSep, "").replace(decimalSep, ".");
            const n = parseFloat(stripped);
            return isNaN(n) ? 0 : n;
        },
        parseUserInput: (val: string): number => {
            if (!val || val.trim() === "") return NaN;
            // Strip spaces (\u00A0 = Finnish/French/Swedish) and apostrophes (Swiss) — these are thousands separators only
            const s = val.trim().replace(/[\s\u00A0']/g, "");
            const hasDot = s.includes(".");
            const hasComma = s.includes(",");

            if (hasDot && hasComma) {
                const lastDot = s.lastIndexOf(".");
                const lastComma = s.lastIndexOf(",");
                if (lastDot > lastComma) {
                    // dot is decimal: "1,000.53"
                    return parseFloat(s.replaceAll(",", ""));
                } else {
                    // comma is decimal: "1.000,53"
                    return parseFloat(s.replaceAll(".", "").replace(",", "."));
                }
            }
            if (hasDot) {
                // Only dots — check if every segment after splitting by "." is numeric and the last is exactly 3 digits
                const segments = s.split(".");
                const lastSeg = segments[segments.length - 1];
                if (/^\d{3}$/.test(lastSeg) && segments.slice(0, -1).every(seg => /^\d+$/.test(seg))) {
                    // German thousands: "1.000" → 1000, "1.000.000" → 1000000
                    return parseFloat(s.replaceAll(".", ""));
                } else {
                    // Decimal: "0.53", "1.5"
                    return parseFloat(s);
                }
            }
            if (hasComma) {
                const segments = s.split(",");
                const lastSeg = segments[segments.length - 1];
                if (/^\d{3}$/.test(lastSeg) && segments.slice(0, -1).every(seg => /^\d+$/.test(seg))) {
                    // en-US thousands: "1,000" → 1000
                    return parseFloat(s.replaceAll(",", ""));
                } else {
                    // German decimal: "0,53"
                    return parseFloat(s.replace(",", "."));
                }
            }
            return parseFloat(s);
        }
    };

    return cachedFormatters;
}