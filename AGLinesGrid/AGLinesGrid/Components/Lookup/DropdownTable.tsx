import * as React from "react";
import {
    TableBody,
    TableCell,
    TableRow,
    Table,
    TableHeader,
    TableHeaderCell,
    useTableFeatures,
    TableColumnDefinition,
    TableColumnId,
    useTableSort,
    createTableColumn,
    makeStyles,
    useTableColumnSizing_unstable,
    Text,
    Spinner
} from "@fluentui/react-components";
import { Grid20Regular } from "@fluentui/react-icons";
import * as lodash from 'lodash';

/**
 * Styles for the control using Fluent UI's makeStyles
 * Defines the layout and appearance of the control components
 */
const useStyles = makeStyles({
    tableHeaderCell: {
        fontWeight: "bold"
    },
    tableCellContent: {
        overflow: "hidden",
        textOverflow: "ellipsis"
    },
    tableBody: {
        height: "300px",
        overflowY: "auto"
    },
    defaultText: {
        display: "flex",
        justifyContent: "center",
        alignItems: "center",
        height: "100%",
        width: "100%"
    }
});

/**
 * DropdownTable JSX element properties interface.
 */
interface DropdownTableProps {
    /** Array of items to display in the table rows. */
    items: any[];
    /** Array of column definitions with `key` and `label` properties. */
    columns: { key: string; label: string }[];
    /** Data loading flag. */
    loading: boolean;
    /** Width of the table. */
    width: number;
    /** Callback function that is triggered when a row is clicked, with the clicked row's item as an argument. */
    onRowClick: (row: any) => void;
    /** Get Localized control strings. */
    getString(id: string): string;
    /** Scroll container ref */
    scrollContainerRef: React.MutableRefObject<HTMLDivElement>
    /** Handles infinite scroll functionality */
    handleScroll: lodash.DebouncedFunc<(e: React.UIEvent<HTMLDivElement> | Event) => Promise<void>>
}

/**
 * DropdownTable component renders a sortable table with dropdown-like behavior.
 * It receives the table data (`items`), the columns metadata (`columns`), and a callback function (`onRowClick`)
 * that gets called when a row is clicked.
 */
export const DropdownTable = (props: DropdownTableProps) => {
    const styles = useStyles();

    const columnsDefsInitial = React.useMemo(
        () =>
            props.columns.map(col =>
                createTableColumn<any>({
                    columnId: col.key,
                    compare: (a, b) => {
                        return a[col.key].localeCompare(b[col.key]);
                    }
                })
            ),
        [props.columns]
    );

    const [columnDefs] = React.useState<TableColumnDefinition<any>[]>(columnsDefsInitial);

    const {
        getRows,
        sort: { getSortDirection, toggleColumnSort, sort },
        columnSizing_unstable,
        tableRef
    } = useTableFeatures(
        {
            columns: columnDefs,
            items: props.items
        },
        [
            useTableSort({
                defaultSortState: { sortColumn: props.columns[0].key, sortDirection: "ascending" }
            }),
            useTableColumnSizing_unstable()
        ]
    );

    const headerSortProps = (columnId: TableColumnId) => ({
        onClick: (e: React.MouseEvent) => {
            toggleColumnSort(e, columnId);
        },
        sortDirection: getSortDirection(columnId)
    });

    const rows = sort(getRows());

    return (
        <div style={{ width: props.width - 15 }}>
            <Table
                ref={tableRef}
                sortable
                noNativeElements
                aria-label="Dropdown table"
                {...columnSizing_unstable.getTableProps()}
            >
                <TableHeader>
                    <TableRow>
                        {props.columns.map(col => (
                            <TableHeaderCell
                                key={col.key}
                                className={styles.tableHeaderCell}
                                {...headerSortProps(col.key)}
                                {...columnSizing_unstable.getTableHeaderCellProps(col.key)}
                            >
                                {col.label}
                            </TableHeaderCell>
                        ))}
                    </TableRow>
                </TableHeader>

                <TableBody className={styles.tableBody} ref={props.scrollContainerRef}>
                    {!props.loading &&
                        rows.length > 0 &&
                        rows.map(({ item }) => (
                            <TableRow key={item.id}>
                                {props.columns.map(col => (
                                    <TableCell
                                        key={item.id + col.key}
                                        onClick={() => props.onRowClick(item)}
                                        {...columnSizing_unstable.getTableCellProps(col.key)}
                                    >
                                        <div className={styles.tableCellContent}>{item[col.key]}</div>
                                    </TableCell>
                                ))}
                            </TableRow>
                        ))}

                    {!props.loading && rows.length === 0 && (
                        <div className={styles.defaultText}>
                            <Grid20Regular />
                            <Text>{props.getString("NoRecordsFound")}</Text>
                        </div>
                    )}

                    {props.loading && (
                        <div className={styles.defaultText}>
                            <Spinner labelPosition="after" label={props.getString("Loading")} />
                        </div>
                    )}
                </TableBody>
            </Table>
        </div>
    );
};
