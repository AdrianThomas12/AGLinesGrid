import * as React from "react";
import "./App.css";
import {
    OverlayDrawer,
    DrawerBody,
    DrawerHeader,
    DrawerHeaderTitle,
    Button,
    useRestoreFocusSource,
    DrawerFooter,
    Field,
    Input,
    Subtitle2
} from "@fluentui/react-components";
import { AddRegular, DeleteRegular, Dismiss24Regular } from "@fluentui/react-icons";
import { AgGridReact } from "ag-grid-react";
import { GridApi, GridReadyEvent, RowEditingStoppedEvent, themeQuartz } from "ag-grid-community";
import { ColDef } from "ag-grid-community";
import { ColumnDefinition } from "../Scripts/types";
import { ConfirmationDialog } from "./ConfirmationDialog";
import { getSplitDeliveryData } from "../Scripts/webApi";
import { buildFormatters } from "../Scripts/locale";

export interface LinesDrawerProps {
    open: boolean;
    headerId: string;
    existing: boolean;
    setOpen: React.Dispatch<React.SetStateAction<boolean>>;
    title: string;
    data: any;
    handleAction: (data: any) => void;
    getString(s: string): string;
    webAPI: ComponentFramework.WebApi;
    debugMode?: boolean;
}

export const SplitDeliveryDrawer = (props: LinesDrawerProps) => {
    const restoreFocusSourceAttributes = useRestoreFocusSource();

    const [rowData, setRowData] = React.useState<any[]>([]);
    const { formatDate, formatDecimal, parseDecimal } = buildFormatters();
    const [totalQuantity, setTotalQuantity] = React.useState<number>(parseFloat(String(props.data?.mserp_purchqty ?? 0)));
    const originalTotalQuantity = React.useRef<number>(parseFloat(String(props.data?.mserp_purchqty ?? 0)));
    const [remainingQuantity, setRemainingQuantity] = React.useState<number>(0);
    const [confirmDialogOpen, setConfirmDialogOpen] = React.useState<boolean>(false);
    const [confirmDeleteLineDialogOpen, setConfirmDeleteLineDialogOpen] = React.useState<boolean>(false);

    const gridRef = React.useRef<AgGridReact>(null);
    const gridContainerRef = React.useRef<HTMLDivElement>(null);
    const drawerRef = React.useRef<HTMLDivElement>(null);
    const gridApiRef = React.useRef<GridApi | null>(null);

    const gridTheme = React.useMemo<any>(
        () =>
            themeQuartz.withParams({
                headerBackgroundColor: "#FFFFFF",
                headerColumnResizeHandleColor: "transparent",
                columnBorder: "solid 2.5px var(--eg-border-color)",
                rangeSelectionBorderColor: "var(--eg-border-color)",
                rangeSelectionBorderStyle: "solid"
            }),
        []
    );

    const columnDefs = React.useMemo<(ColDef & ColumnDefinition)[]>(
        () => [
            {
                headerName: "",
                headerCheckboxSelection: true,
                checkboxSelection: true,
                width: 50,
                sortable: false
            },
            {
                headerName: "Quantity",
                field: "mserp_purchqty",
                width: 250,
                cellEditor: "agNumberCellEditor",
                editable: true,
                valueParser: params => {
                    const newValue = parseDecimal(params.newValue);
                    return newValue === 0 && params.oldValue ? params.oldValue : newValue;
                },
                valueFormatter: params => {
                    const v = typeof params.value === "string" ? Number(params.value) : params.value;
                    if (v == null || isNaN(v)) return "";
                    return formatDecimal(v, 2);
                },
                valueSetter: params => {
                    if (!params.data) return false;

                    params.data.mserp_purchqty = params.newValue;
                    return true;
                }
            },
            {
                headerName: "Requested Receipt Date",
                field: "mserp_deliverydate",
                width: 250,
                cellEditor: "agDateCellEditor",
                editable: false,
                valueFormatter: params => {
                    const value = params.value as Date;
                    return value instanceof Date && !isNaN(value.getTime()) ? formatDate(value) : "";
                }
            },
            {
                headerName: "Confirmed Receipt Date",
                field: "mserp_confirmeddlv",
                width: 250,
                cellEditor: "agDateCellEditor",
                editable: true,
                valueFormatter: params => {
                    const value = params.value as Date;
                    return value instanceof Date && !isNaN(value.getTime()) ? formatDate(value) : "";
                },
                valueSetter: params => {
                    if (!params.data) return false;

                    params.data.mserp_confirmeddlv = params.newValue;
                    return true;
                }
            }
        ],
        []
    );

    /**
     * Adds a new row into the AG Grid
     */
    const addNewRow = () => {
        const newItem = {
            id: rowData.length + 1,
            mserp_purchqty: 0.0,
            mserp_deliverydate: props.data.mserp_deliverydate,
            mserp_confirmeddlv: props.data.mserp_confirmeddlv
        };

        gridRef.current!.api.applyTransaction({ add: [newItem] });
        setRowData(prevRows => [...prevRows, newItem]);
    };

    /**
     * Deletes the selected rows from the AG Grid
     */
    const deleteSelectedRows = () => {
        if (!gridRef.current) return;

        const selectedNodes = gridRef.current.api.getSelectedNodes();
        if (selectedNodes.length === 0) return;

        const selectedIds = new Set(selectedNodes.map(node => node.data.id));

        gridRef.current.api.applyTransaction({
            remove: selectedNodes.map(node => node.data)
        });

        let sumOfConfirmedQuantities = 0;
        gridRef.current.api.forEachNode(row => {
            sumOfConfirmedQuantities += parseFloat(String(row?.data?.mserp_purchqty ?? 0));
        });

        setRowData(prevRows => prevRows.filter(row => !selectedIds.has(row.id)));
        setTotalQuantity(sumOfConfirmedQuantities);
        setRemainingQuantity(originalTotalQuantity.current - sumOfConfirmedQuantities);
        setConfirmDeleteLineDialogOpen(false);
    };

    /**
     * Handles calculation on the remaining quantity based on the quantities in the AG Grid
     * @param event RowEditingStopped Event
     */
    const handleRowEditingStopped = (event: RowEditingStoppedEvent<any>) => {
        if (!gridRef.current) return;

        if (event.event instanceof KeyboardEvent && event.event.key.toLowerCase() === "escape") {
            event.node.setSelected(false);
            return;
        }

        let sumOfConfirmedQuantities = 0;
        gridRef.current.api.forEachNode(row => {
            sumOfConfirmedQuantities += parseFloat(String(row?.data?.mserp_purchqty ?? 0));
        });

        const updatedRows = rowData.map(row => {
            if (row.id === event.data.id) {
                return {
                    ...row,
                    mserp_purchqty: event.data.mserp_purchqty ?? row.mserp_purchqty,
                    mserp_confirmeddlv: event.data.mserp_confirmeddlv ?? row.mserp_confirmeddlv
                };
            }
            return row;
        });

        setRowData(updatedRows);
        setTotalQuantity(sumOfConfirmedQuantities);
        setRemainingQuantity(originalTotalQuantity.current - sumOfConfirmedQuantities);
        if (gridRef.current) gridRef.current.api.ensureIndexVisible(event.rowIndex, "middle");
    };

    /**
     * Data transformation prepared to sent to the action request.
     * Opens confirmation dialog when the total quantity is not the same
     * as the original total quantity.
     *
     * Sends input parameters of type LineSplitDeliveryAction
     */
    const handleDataAndAction = () => {
        if (remainingQuantity !== 0) {
            setConfirmDialogOpen(true);
        } else {
            const params: string = setSplitDeliveryParams();
            props.handleAction(params);
        }
    };

    /**
     * Opens confirmation dialog for when the
     * total quantity is not the same as the original quantity
     */
    const onConfirmationDialogConfirm = () => {
        const params: string = setSplitDeliveryParams();
        props.handleAction(params);
    };

    /**
     * Data transformation for the action parameters.
     * @returns {LineSplitDeliveryAction} Input parameters array for the split delivery OData action
     *
     * The Value needs to be a string as FO does not recognize json objects.
     *
     * @description How a simple object into the string array will look like
     * @example { Quantity: 10.00, ConfirmedReceiptDate: '2025-04-10' }
     */
    const setSplitDeliveryParams = (): string => {
        return JSON.stringify([
            {
                Name: "_deliveries",
                Value: `[${rowData
                    .map(
                        row =>
                            `{ Quantity: ${row.mserp_purchqty ?? "0.00"}, ConfirmedReceiptDate: '${
                                row.mserp_confirmeddlv instanceof Date
                                    ? `${row.mserp_confirmeddlv.getFullYear()}-${String(
                                          row.mserp_confirmeddlv.getMonth() + 1
                                      ).padStart(2, "0")}-${String(row.mserp_confirmeddlv.getDate()).padStart(2, "0")}`
                                    : row.mserp_confirmeddlv
                            }' }`
                    )
                    .join(", ")}]`.trim()
            }
        ]);
    };

    /**
     * Sets the grid data
     */
    const onGridReady = (event: GridReadyEvent<any>) => {
        gridApiRef.current = event.api;

        if (props.existing) return;

        setRowData([
            {
                mserp_purchqty: props?.data?.mserp_purchqty,
                mserp_deliverydate: props?.data?.mserp_deliverydate,
                mserp_confirmeddlv: props?.data?.mserp_confirmeddlv
            }
        ]);
    };

    const handleCellClick = (event: any) => {
        const colKey = event.column.getColId();
        const rowIndex = event.rowIndex;

        if (event.node.isSelected()) {
            event.api.startEditingCell({
                rowIndex,
                colKey
            });
        }
    };

    React.useEffect(() => {
        if (!props.existing) {
            setRowData([
                {
                    mserp_purchqty: props?.data?.mserp_purchqty,
                    mserp_deliverydate: props?.data?.mserp_deliverydate,
                    mserp_confirmeddlv: props?.data?.mserp_confirmeddlv
                }
            ]);
            originalTotalQuantity.current = parseFloat(String(props?.data?.mserp_purchqty ?? 0));
            setTotalQuantity(parseFloat(String(props?.data?.mserp_purchqty ?? 0)));
        }
    }, [props.data]);

    React.useEffect(() => {
        if (!props.open) return;

        const loadData = async () => {
            const qty = parseFloat(String(props.data?.mserp_purchqty ?? 0));
            originalTotalQuantity.current = qty;
            setTotalQuantity(qty);
            setRemainingQuantity(0);

            if (props.existing) {
                try {
                    const existingRowData = await getSplitDeliveryData(
                        props.headerId,
                        props?.data?.mserp_linenumber,
                        props.webAPI
                    );
                    setRowData(existingRowData);
                } catch (err) {
                    console.error("Failed to load split delivery data", err);
                    alert(err);
                }
            } else {
                const defaultRow = {
                    mserp_purchqty: props?.data?.mserp_purchqty,
                    mserp_deliverydate: props?.data?.mserp_deliverydate,
                    mserp_confirmeddlv: props?.data?.mserp_confirmeddlv
                };
                setRowData([defaultRow]);
                setConfirmDialogOpen(false);
            }
        };

        loadData();
    }, [props.open]);

    const stopEditingAndClearSelection = React.useCallback(() => {
        const api = gridApiRef.current;
        if (!api) return;
        if (api.getEditingCells()?.length) api.stopEditing();
        if (api.getSelectedNodes()?.length) api.deselectAll();
    }, []);

    React.useEffect(() => {
        /**
         * Handles clicking outside of the Grid and anywhere on the page.
         * If the row is in edit, it will stop editing and save the record
         * @param event Mouse event
         */
        const handleClickOutside = (event: MouseEvent) => {
            const container = gridContainerRef.current;
            const target = event.target as HTMLElement | null;
            if (!container || !target) return;

            const clickedToolbarAction =
                !!target.closest("#delete-checkbox") ||
                !!target.closest("[data-grid-action]");

            // if click is outside the grid container → stop edit & deselect
            if (!container.contains(target) && !clickedToolbarAction) {
                stopEditingAndClearSelection();
            }
        };

        document.addEventListener("mousedown", handleClickOutside);

        return () => {
            document.removeEventListener("mousedown", handleClickOutside);
        };
    });

    return (
        <OverlayDrawer
            size={"large"}
            {...restoreFocusSourceAttributes}
            position="end"
            open={props.open}
            onOpenChange={(_, state) => props.setOpen(state.open)}
            modalType="alert"
        >
            <DrawerHeader>
                <DrawerHeaderTitle
                    action={
                        <Button
                            appearance="subtle"
                            aria-label="Close"
                            icon={<Dismiss24Regular />}
                            onClick={() => props.setOpen(false)}
                        />
                    }
                >
                    {props.title}
                </DrawerHeaderTitle>
            </DrawerHeader>

            <DrawerBody>
                <div ref={drawerRef}>
                    <div className="divider-container">
                        <Subtitle2 className="divider-container-title">{props.getString("Summary")}</Subtitle2>
                    </div>

                    <div className="row">
                        <div className="col-lg-4 col-md-4 col-sm-12">
                            <Field label={props.getString("ItemNumber")}>
                                <Input readOnly disabled value={props.data?.mserp_itemid} />
                            </Field>
                        </div>

                        <div className="col-lg-4 col-md-4 col-sm-12">
                            <Field label={props.getString("ProcurementCategory")}>
                                <Input readOnly disabled value={props.data?.mserp_procurementcategory} />
                            </Field>
                        </div>

                        <div className="col-lg-4 col-md-4 col-sm-12">
                            <Field label={props.getString("TotalQuantity")}>
                                <Input readOnly disabled value={parseFloat(totalQuantity.toString()).toFixed(2)} />
                            </Field>
                            <Field label={props.getString("RemainingQuantity")}>
                                <Input readOnly disabled value={parseFloat(remainingQuantity.toString()).toFixed(2)} />
                            </Field>
                        </div>
                    </div>

                    {props.debugMode && <div className="eg-debug-panel">
                        <div className="eg-debug-panel__header">DRAWER DEBUG v1.2</div>
                        <div className="eg-debug-panel__body">
                            <table className="eg-debug-panel__table">
                                <thead>
                                    <tr>
                                        <th className="eg-debug-panel__th">Source</th>
                                        <th className="eg-debug-panel__th">mserp_purchqty (type)</th>
                                        <th className="eg-debug-panel__th">raw API?</th>
                                        <th className="eg-debug-panel__th">parseFloat</th>
                                        <th className="eg-debug-panel__th">parseDecimal</th>
                                        <th className="eg-debug-panel__th">totalQuantity state</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <tr>
                                        <td className="eg-debug-panel__td">props.data</td>
                                        <td className="eg-debug-panel__td"><code>"{String(props.data?.mserp_purchqty)}" ({typeof props.data?.mserp_purchqty})</code></td>
                                        <td className="eg-debug-panel__td"><code>{/^\d+\.\d+$/.test(String(props.data?.mserp_purchqty)) ? "YES" : "NO"}</code></td>
                                        <td className="eg-debug-panel__td"><code>{parseFloat(String(props.data?.mserp_purchqty))}</code></td>
                                        <td className="eg-debug-panel__td"><code>{parseDecimal(props.data?.mserp_purchqty)}</code></td>
                                        <td className="eg-debug-panel__td"><code>{totalQuantity}</code></td>
                                    </tr>
                                    {rowData.map((row, i) => (
                                        <tr key={i}>
                                            <td className="eg-debug-panel__td">rowData[{i}]</td>
                                            <td className="eg-debug-panel__td"><code>"{String(row.mserp_purchqty)}" ({typeof row.mserp_purchqty})</code></td>
                                            <td className="eg-debug-panel__td"><code>{/^\d+\.\d+$/.test(String(row.mserp_purchqty)) ? "YES" : "NO"}</code></td>
                                            <td className="eg-debug-panel__td"><code>{parseFloat(String(row.mserp_purchqty))}</code></td>
                                            <td className="eg-debug-panel__td"><code>{parseDecimal(row.mserp_purchqty)}</code></td>
                                            <td className="eg-debug-panel__td">—</td>
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                        </div>
                    </div>}

                    <div className="divider-container">
                        <Subtitle2 className="divider-container-title">
                            {props.getString("DeliveryForPurchaseLine")}
                        </Subtitle2>
                    </div>

                    <Button appearance="subtle" size="small" icon={<AddRegular />} onClick={addNewRow}>
                        {props.getString("New")}
                    </Button>
                    <Button
                        id="delete-checkbox"
                        data-grid-action="delete"
                        appearance="subtle"
                        size="small"
                        icon={<DeleteRegular />}
                        onClick={() => {
                            if (props.existing) {
                                setConfirmDeleteLineDialogOpen(true);
                            } else {
                                deleteSelectedRows();
                            }
                        }}
                    >
                        {props.getString("Delete")}
                    </Button>

                    <div id="split-delivery-grid" ref={gridContainerRef}>
                        <AgGridReact
                            ref={gridRef}
                            theme={gridTheme}
                            rowData={rowData}
                            columnDefs={columnDefs}
                            singleClickEdit={true}
                            suppressPaginationPanel
                            paginationPageSizeSelector={false}
                            rowSelection="multiple"
                            editType="fullRow"
                            onGridReady={onGridReady}
                            domLayout="normal"
                            onRowEditingStopped={handleRowEditingStopped}
                            onCellClicked={handleCellClick}
                        />
                    </div>
                </div>
            </DrawerBody>

            <DrawerFooter>
                <Button appearance="primary" onClick={handleDataAndAction}>
                    {props.getString("OK")}
                </Button>
                <Button onClick={() => props.setOpen(false)}>{props.getString("Cancel")}</Button>
            </DrawerFooter>

            <ConfirmationDialog
                isOpen={confirmDialogOpen}
                title={props.getString("SplitDeliveryQuantityTitle")}
                content={props.getString("SplitDeliveryQuantityMessage")}
                onConfirm={onConfirmationDialogConfirm}
                onClose={() => setConfirmDialogOpen(false)}
                getString={props.getString}
            />

            <ConfirmationDialog
                isOpen={confirmDeleteLineDialogOpen}
                title={props.getString("DeleteRecordTitle")}
                content={props.getString("DeleteSplitDeliveryConfirmation")}
                onConfirm={deleteSelectedRows}
                onClose={() => setConfirmDeleteLineDialogOpen(false)}
                getString={props.getString}
            />
        </OverlayDrawer>
    );
};