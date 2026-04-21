/**
 * @file GridSortableHeader.tsx
 * @description Component for rendering sortable column headers in the grid.
 */

import * as React from "react";
import { CustomHeaderProps } from "ag-grid-react";

import "./App.css";

/**
 * Component that renders a sortable column header for the grid.
 *
 * @param {CustomHeaderProps & React.PropsWithChildren} props - The properties for the sortable header.
 * @returns {JSX.Element} The rendered sortable header component.
 */
export const GridSortableHeader = (props: React.PropsWithChildren<CustomHeaderProps>) => {
    const [sortType, setSortType] = React.useState<any>(undefined);

    /**
     * Updates the sort state when sorting changes.
     */
    const onSortChanged = () => {
        const sort = props.column.getSort();

        setSortType(sort);
    };

    /**
     * Requests sorting based on the current state and user interaction.
     *
     * @param {'asc' | 'desc' | null} order - The sorting order.
     * @param {any} event - The event that triggered sorting.
     */
    const onSortRequested = (order: "asc" | "desc" | null, event: any) => {
        if (!props.enableSorting) {
            return;
        }

        const newOrder = order === "asc" ? "desc" : order === "desc" ? null : "asc";

        props.setSort(newOrder, event.shiftKey);
    };

    React.useEffect(() => {
        props.column.addEventListener("sortChanged", onSortChanged);
        onSortChanged();

        return () => {
            props.column.removeEventListener("sortChanged", onSortChanged);
        };
    }, []);

    return (
        <div className="eg-sortable-header" onClick={event => onSortRequested(sortType, event)}>
            {props.children}
            {sortType === "asc" ? (
                <span className="ag-sort-indicator-icon ag-sort-ascending-icon" aria-hidden="true">
                    <span className="ag-icon ag-icon-asc" unselectable="on" role="presentation"></span>
                </span>
            ) : sortType === "desc" ? (
                <span className="ag-sort-indicator-icon ag-sort-descending-icon" aria-hidden="true">
                    <span className="ag-icon ag-icon-desc" unselectable="on" role="presentation"></span>
                </span>
            ) : null}
        </div>
    );
};
