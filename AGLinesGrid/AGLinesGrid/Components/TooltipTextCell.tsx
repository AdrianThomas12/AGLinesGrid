/**
 * @file TooltipTextCell.tsx
 * @description Component for rendering a text cell with a tooltip when content overflows.
 */

import * as React from "react";
import "./App.css";
import { Tooltip } from "@fluentui/react-components";
import { buildFormatters } from "../Scripts/locale";

/**
 * Component that renders a text cell with an optional tooltip.
 *
 * @param {object} props - The component props.
 * @param {any} props.value - The main value displayed in the cell.
 * @param {any} [props.valueFormatted] - The formatted value for display.
 * @param {any} [props.valueBefore] - The original value before any modifications.
 * @returns {JSX.Element} The rendered tooltip text cell component.
 */
export const TooltipTextCell = (props: {
    value: any;
    originalValueText?: string;
    valueFormatted?: any;
    valueBefore?: any;
    dataType?: string;
}) => {
    const { formatDecimal, formatDate } = buildFormatters();

    const formatBefore = (val: any): string => {
        if (val instanceof Date) return formatDate(val);
        if (
            props.dataType === "decimal" ||
            props.dataType === "double" ||
            props.dataType === "money"
        ) return formatDecimal(parseFloat(String(val ?? "")), 2);
        return String(val ?? "");
    };

    const [overflowActive, setOverflowActive] = React.useState<boolean>(false);
    const [visible, setVisible] = React.useState(false);
    const [anchorRef, setAnchorRef] = React.useState<HTMLSpanElement | null>(null);

    const value = React.useMemo(() => {
        if (props.value instanceof Date) return props.valueFormatted;
        return props.valueFormatted ?? props.value;
    }, [props.valueFormatted, props.value]);

    const showOriginalValue = React.useMemo(() => {
        if (props.valueBefore && props.valueBefore instanceof Date && props.value instanceof Date)
            return props.value.getTime() !== props.valueBefore.getTime();

        if (
            props.dataType === "decimal" ||
            props.dataType === "double" ||
            props.dataType === "integer" ||
            props.dataType === "money"
        )
            return (
                props.valueBefore &&
                parseFloat(String(props.valueBefore)) !== parseFloat(String(props.value))
            );

        return props.valueBefore && props.value !== props.valueBefore;
    }, [props.valueBefore, props.value]);

    /**
     * Handles mouse enter events to determine if the text content is overflowing.
     * If the content overflows, it enables the tooltip display.
     *
     * @param {React.MouseEvent<HTMLDivElement, MouseEvent>} event - The mouse enter event.
     */
    const handle_Mouse_Enter = React.useCallback((event: React.MouseEvent<HTMLDivElement, MouseEvent>) => {
        const e = event.currentTarget;
        const ellipsis = e.offsetHeight < e.scrollHeight || e.offsetWidth < e.scrollWidth;

        setOverflowActive(ellipsis);
    }, []);

    return (
        <Tooltip
            content={
                showOriginalValue
                    ? `${props.originalValueText}: ${formatBefore(props.valueBefore)}`
                    : value
            }
            withArrow
            positioning={{ target: overflowActive ? undefined : anchorRef, position: "above" }}
            showDelay={500}
            hideDelay={200}
            relationship="description"
            visible={visible && (overflowActive || showOriginalValue)}
            onVisibleChange={(_ev, data) => setVisible(data.visible)}
        >
            <div className="eg-overflow-text" onMouseEnter={handle_Mouse_Enter}>
                {showOriginalValue && <span className="eg-green-dot">●</span>}
                <span ref={setAnchorRef}>{value}</span>
            </div>
        </Tooltip>
    );
};