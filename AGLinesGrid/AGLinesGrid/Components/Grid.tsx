/**
 * @file Grid.tsx
 * @description This file contains the Grid component, which integrates with ag-Grid to provide a powerful and customizable data table.
 * The Grid supports features such as sorting, filtering, column resizing, custom cell rendering, and row actions.
 */

import * as React from "react";

import type {
    ColDef,
    EditableCallbackParams,
    GridApi,
    RowClassParams,
    RowEditingStartedEvent,
    RowEditingStoppedEvent,
    RowSelectedEvent,
    ValueFormatterParams
} from "ag-grid-community";
import { ModuleRegistry } from "@ag-grid-community/core";
import { ClientSideRowModelModule } from "@ag-grid-community/client-side-row-model";
import { themeQuartz } from "ag-grid-community";
import { AgGridReact, CustomCellRendererProps } from "ag-grid-react";
import {
    AttachmentEntityName,
    AttachmentFieldName,
    ColumnDefinition,
    FieldNames,
    GridProps,
    GridRowActionProps,
    LineType,
    NoYes,
    PCFCommunication,
    PCFCommunicationType,
    VirtualEntityActions
} from "../Scripts/types";
import { GridRowActions } from "./GridRowActions";
import {
    SearchBoxChangeEvent,
    InputOnChangeData,
    SearchBox,
    Switch,
    SwitchOnChangeData,
    Spinner,
    FluentProvider,
    webLightTheme,
    Button
} from "@fluentui/react-components";
import { GridSortableHeader } from "./GridSortableHeader";

import "./App.css";
import { TooltipTextCell } from "./TooltipTextCell";
import { cloneDeep } from "lodash";
import { CustomPagination } from "./Pagination";
import { AddRegular, AttachRegular, CommentMultipleRegular } from "@fluentui/react-icons";
import { _getPlural, executeAction, getFetchXml, getLookupValues, normalizeCheckboxValue } from "../Scripts/webApi";
import { LookupCellEditor } from "./Lookup/LookupCellEditor";

import { LookupProps } from "./Lookup/Lookup";
import { DetailsPopover } from "./DetailsPopover";
import { ConfirmationDialog } from "./ConfirmationDialog";
import { buildFormatters } from "../Scripts/locale";

ModuleRegistry.registerModules([ClientSideRowModelModule]);

/**
 * Debounces a function call to limit execution frequency.
 *
 * @param {any} fn - The function to debounce.
 * @param {number} ms - The debounce timeout in milliseconds.
 * @returns {any} A debounced function.
 */
function debounce(this: any, fn: any, ms: any): any {
    let timer: any;
    return () => {
        clearTimeout(timer);
        timer = setTimeout(function () {
            timer = null;
            fn.apply(this, arguments);
        }, ms);
    };
}

/**
 * Merges default column definitions with user-defined custom columns.
 *
 * @param {ColumnDefinition[]} columns - The default column definitions.
 * @param {ColumnDefinition[]} customColumns - The user-defined custom columns.
 * @returns {(ColumnDefinition)[]} The merged column definitions.
 */
function customizeColumns(columns: ColumnDefinition[], customColumns: ColumnDefinition[]): ColumnDefinition[] {
    const resultColumns = [...columns];

    if (!customColumns) return resultColumns;

    for (const customCol of customColumns) {
        const i = resultColumns.findIndex(c => c.field === customCol.field);

        if (i !== -1) {
            resultColumns[i] = { ...resultColumns[i], ...customCol };
        } else {
            // Not found, just push it as a new column (e.g., extras like "comments")
            resultColumns.push(customCol);
        }
    }

    return resultColumns;
}

/**
 * Configures grid column properties based on customization options and data types.
 *
 * @param {ColumnDefinition[]} columnDefs - The grid column definitions.
 * @param {any[] | null} retrievedRowData - The grid row data.
 * @param {boolean} readonly - Whether the grid is in read-only mode.
 * @returns {(ColDef & ColumnDefinition)[]} The configured column definitions for ag-Grid.
 */
function getGridColumns(
    columnDefs: ColumnDefinition[],
    retrievedRowData: any[] | null,
    readonly: boolean,
    lineStatusLogicalName: string,
    attributesMetadata: any[],
    getString: (s: string) => string,
    webApi: ComponentFramework.WebApi,
    hidePriceField: string | undefined,
    hidePriceFields: string[],
    vendorId: string | null,
    isReadOnlyInvoice: boolean
): (ColDef & ColumnDefinition)[] {
    const { formatDate, formatInteger, formatDecimal, parseDecimal, parseUserInput } = buildFormatters();
    const gridColumns = columnDefs.map(
        (cd, i) =>
            ({
                ...cd,
                hide:
                    cd.toggleShowColumnEnabled === false ||
                    (hidePriceField &&
                        normalizeCheckboxValue(retrievedRowData?.find(d => d[hidePriceField])?.[hidePriceField]) ===
                            NoYes.No &&
                        hidePriceFields?.includes(cd.field)),
                sortable: cd.sortable === false ? false : true,
                fieldOriginalDisplayName: cd.fieldOriginalValue
                    ? `${
                          attributesMetadata.find(m => m.LogicalName === cd.fieldOriginalValue)?.DisplayName
                              ?.UserLocalizedLabel?.Label
                      }:`
                    : getString("OriginalValue"),
                editable: (params: EditableCallbackParams<any>) => {
                    if (isReadOnlyInvoice) {
                        return false;
                    }

                    const isRejectedOrSubstitutedOrSplit =
                        params.data?.[lineStatusLogicalName] === "Rejected" ||
                        params.data?.[lineStatusLogicalName] === "Substituted" ||
                        params.data?.[lineStatusLogicalName] === "Split into schedule";

                    const isEditableByStatus =
                        cd.editableOnStatus && lineStatusLogicalName
                            ? cd.editableOnStatus.includes(params.data?.[lineStatusLogicalName]) &&
                              !isRejectedOrSubstitutedOrSplit &&
                              cd.allowUpdate
                            : !isRejectedOrSubstitutedOrSplit && cd.allowUpdate;

                    const isAlternateLineEditable =
                        cd.editableOnAlternate &&
                        cd.alternateField &&
                        cd.editableAlternateFields &&
                        Array.isArray(cd.editableAlternateFields) &&
                        normalizeCheckboxValue(params.data?.[cd.alternateField]) === NoYes.Yes &&
                        cd.editableAlternateFields.includes(cd.field);

                    const isParentLine =
                        params.data?.hasOwnProperty("mserp_parentline") &&
                        normalizeCheckboxValue(params.data?.mserp_parentline) === NoYes.Yes
                            ? true
                            : false;

                    if (isParentLine) {
                        return false;
                    }

                    if (!readonly && isEditableByStatus) {
                        return true;
                    }

                    return !readonly && (cd.editable || (isAlternateLineEditable && cd.editableOnAlternate));
                },
                cellRenderer: cd.showAsCheckbox
                    ? (props: any) => {
                          const isChecked = normalizeCheckboxValue(props.value) === NoYes.Yes;
                          return (
                              <input
                                  type="checkbox"
                                  disabled
                                  checked={isChecked}
                                  className="ag-input-field-input ag-checkbox-input ag-checkbox-custom-input"
                                  name={props.colDef.field}
                              />
                          );
                      }
                    : (props: any) => {
                          return (
                              <TooltipTextCell
                                  value={
                                      typeof props.value === "object"
                                          ? props.value instanceof Date
                                              ? props.value
                                              : props.value?.name ?? ""
                                          : props.value
                                  }
                                  originalValueText={
                                      attributesMetadata.find(m => m.LogicalName === cd.fieldOriginalValue)?.DisplayName
                                          ?.UserLocalizedLabel?.Label ?? getString("OriginalValue")
                                  }
                                  valueFormatted={props.valueFormatted}
                                  valueBefore={
                                      (cd as any).showOriginalValueTooltip &&
                                      retrievedRowData &&
                                      retrievedRowData?.[props.node.rowIndex]?.[
                                          props.colDef.fieldOriginalValue ?? props.colDef.field
                                      ]
                                  }
                                  dataType={cd.dataType}
                              />
                          );
                      },
                cellEditor:
                    cd.dataType === "string"
                        ? "agTextCellEditor"
                        : cd.dataType === "decimal" ||
                          cd.dataType === "double" ||
                          cd.dataType === "integer" ||
                          cd.dataType === "money"
                        ? "agTextCellEditor"
                        : cd.dataType === "picklist" || cd.dataType === "state" || cd.dataType === "status"
                        ? cd.showAsCheckbox === true
                            ? "agCheckboxCellEditor"
                            : "agSelectCellEditor"
                        : cd.dataType === "boolean"
                        ? "agSelectCellEditor"
                        : cd.dataType === "datetime"
                        ? "agDateCellEditor"
                        : cd.dataType === "lookup"
                        ? "lookupComponent"
                        : null,

                cellEditorParams: (params: any) => {
                    return cd.dataType === "lookup"
                        ? ({
                              getValues: (searchText: string, pageNumber: number) =>
                                  getLookupValues(
                                      getFetchXml(
                                          attributesMetadata.find(m => m.LogicalName === cd.field)?.Targets?.[0],
                                          cd.columns[0].key,
                                          null,
                                          null,
                                          null,
                                          searchText,
                                          pageNumber,
                                          false,
                                          vendorId
                                      ),
                                      attributesMetadata.find(m => m.LogicalName === cd.field)?.Targets?.[0],
                                      cd.columns[0].key,
                                      null,
                                      pageNumber,
                                      webApi
                                  ),
                              displayField: cd.columns[0].key,
                              columns: cd.columns,
                              webApi: webApi,
                              getString: getString,
                              configuration: {
                                  lookupEntityType: attributesMetadata.find(m => m.LogicalName === cd.field)
                                      ?.Targets[0],
                                  displayField: cd.columns[0].key,
                                  secondaryDisplayField: null,
                                  filterRecordId: null,
                                  filterFieldRelationColumn: null
                              },
                              selectedValue: params?.data?.[cd.field]
                          } as LookupProps)
                        : {
                              values: cd.options?.map(o => o.label)
                          };
                },
                valueParser:
                    cd.dataType === "decimal" ||
                    cd.dataType === "double" ||
                    cd.dataType === "money" ||
                    cd.dataType === "integer"
                        ? params => {
                              const p = parseUserInput(String(params.newValue ?? ""));
                              return isNaN(p) ? params.oldValue : p;
                          }
                        : undefined,
                valueFormatter:
                    cd.dataType !== "datetime"
                        ? cd.dataType === "decimal" || cd.dataType === "double" || cd.dataType === "money"
                            ? params => {
                                  const v = typeof params.value === "string" ? Number(params.value) : params.value;
                                  if (v == null || isNaN(v)) return "";
                                  return formatDecimal(v, cd.convertToInteger ? 0 : 2);
                              }
                            : cd.dataType === "integer"
                            ? params => {
                                  const v = typeof params.value === "string" ? Number(params.value) : params.value;
                                  return v == null || isNaN(v) ? "" : formatInteger(v);
                              }
                            : cd.dataType === "lookup"
                            ? params => params?.data?.[cd.field]?.name ?? ""
                            : undefined
                        : (params: ValueFormatterParams<any, Date>) => {
                              const value = params.value as Date;
                              return value instanceof Date && !isNaN(value.getTime()) ? formatDate(value) : "";
                          },
                valueGetter: params => {
                    if (cd.dataType === "integer")
                        return typeof params.data?.[cd.field] === "string"
                            ? parseInt(params.data?.[cd.field]).toString()
                            : params.data?.[cd.field].toString();

                    if (cd.dataType === "decimal" || cd.dataType === "double" || cd.dataType === "money")
                        return typeof params.data?.[cd.field] === "string"
                            ? cd.convertToInteger
                                ? parseFloat(params.data?.[cd.field]).toFixed(0)
                                : parseFloat(params.data?.[cd.field]).toFixed(2)
                            : cd.convertToInteger
                            ? params.data?.[cd.field].toFixed(0)
                            : params.data?.[cd.field].toFixed(2);

                    if (cd.dataType === "lookup") {
                        return params.data?.[cd.field] ?? null;
                    }

                    if (cd.showAsCheckbox) {
                        return normalizeCheckboxValue(params.data?.[cd.field]) === NoYes.Yes ? true : false;
                    }

                    return params.data?.[cd.field] ?? "";
                },
                valueSetter: params => {
                    if (cd.dataType === "decimal" || cd.dataType === "double" || cd.dataType === "money") {
                        const rawField = params.data[cd.field];
                        const oldValue = typeof rawField === "number" ? rawField : parseFloat(String(rawField ?? ""));
                        const newValue = typeof params.newValue === "number" ? params.newValue : parseUserInput(String(params.newValue ?? ""));

                        if (!isNaN(newValue) && oldValue !== newValue) {
                            params.data[cd.field] = newValue;
                            return true;
                        }

                        return false;
                    }

                    if (params.oldValue instanceof Date && params.newValue instanceof Date) {
                        if (params.oldValue.getTime() !== params.newValue.getTime()) {
                            params.data[cd.field] = params.newValue;
                            return true;
                        } else return false;
                    }

                    if (cd.dataType === "lookup") {
                        if (params.oldValue !== params.newValue) {
                            params.data[cd.field] = params.newValue;
                            return true;
                        }

                        return false;
                    }

                    if (cd.showAsCheckbox) {
                        const newValue = params.newValue === true ? NoYes.Yes : NoYes.No;
                        const oldValue = normalizeCheckboxValue(params.data[params.colDef.field]);
                        if (newValue !== oldValue) {
                            params.data[params.colDef.field] = newValue;
                            return true;
                        }
                        return false;
                    }

                    if (params.oldValue !== params.newValue) {
                        params.data[cd.field] = params.newValue;
                        return true;
                    }

                    return false;
                },
                equals: cd.dataType === "lookup" ? (a: any, b: any) => a?.id === b?.id : undefined,
                lockPinned: i <= 1,
                suppressMovable: i <= 1,
                cellStyle:
                    cd.dataType === "lookup"
                        ? { display: "flex", alignItems: "center", justifyContent: "center" }
                        : cd.showAsCheckbox
                        ? { display: "flex", justifyContent: "center" }
                        : {},
                cellClass: cd.showAsCheckbox ? "eg-cell-checkbox-center" : undefined
            } as ColDef & ColumnDefinition)
    );

    return gridColumns;
}

/**
 * Resizes the grid columns to fit the container.
 *
 * @param {GridApi | undefined} api - The ag-Grid API.
 * @param {ColDef[] | null} columnDefs - The column definitions.
 */
function resizeColumns(api: GridApi | undefined, columnDefs: ColDef[] | null) {
    if (!api || !columnDefs) return;

    const limits = columnDefs
        .filter(cd => cd.field) // Only include columns with a valid field
        .map(cd => ({
            key: cd.field!,
            minWidth: cd.width
        }));

    api.sizeColumnsToFit({ columnLimits: limits });
}

/**
 * Pins columns up to a specified index in the grid.
 *
 * @param {number} pinIndex - The index up to which columns should be pinned.
 * @param {ColDef[]} gridColumns - The column definitions.
 */
function pinColumns(pinIndex: number, gridColumns: ColDef[]) {
    for (let i = 0; i <= pinIndex; ++i) {
        gridColumns[i].pinned = "left";
    }
}

/**
 * Converts data between formats used in Dataverse and the grid application.
 */
const convertRowData = {
    /**
     * Converts Dataverse data format to the grid format.
     *
     * @param {any} rowData - The row data object.
     * @param {(ColDef & ColumnDefinition)[] | null} colDefs - The column definitions.
     */
    dataverseToGridApp: (rowData: any, colDefs: (ColDef & ColumnDefinition)[] | null) => {
        for (const cd of colDefs ?? []) {
            if (cd.dataType === "datetime") {
                const date = new Date(rowData[cd.field!]);

                rowData[cd.field!] = isNaN(date.getTime()) ? null : date;
            }
        }
    },

    /**
     * Converts grid data format back to Dataverse format.
     *
     * @param {any} rowData - The row data object.
     * @param {(ColDef & ColumnDefinition)[] | null} colDefs - The column definitions.
     */
    gridAppToDataverse: (rowData: any, colDefs: (ColDef & ColumnDefinition)[] | null) => {
        for (const cd of colDefs ?? []) {
            if (cd.options) {
                rowData[cd.field!] = cd.options.find(o => o.label === rowData[cd.field!])?.value ?? null;
            }
        }
    }
};

/** Custom loader with Fluent UI spinner */
const CustomLoadingOverlay = () => {
    return <Spinner labelPosition="after" label="Loading..." />;
};

/**
 * The main Grid component, providing an interactive table with custom column rendering, sorting, filtering, and row actions.
 *
 * @component
 * @param {GridProps} props - The properties passed to the grid.
 * @returns {JSX.Element} The rendered grid component.
 */
export const Grid = (props: GridProps) => {
    const { parseDecimal, formatDecimal, formatInteger, parseUserInput } = buildFormatters();
    const [retrievedColDefs, setRetrievedColDefs] = React.useState<ColumnDefinition[] | null>(null);
    const [retrievedRowData, setRetrievedRowData] = React.useState<any[] | null>(null);
    const [colDefs, setColDefs] = React.useState<(ColDef & ColumnDefinition)[] | null>([]);
    const [rowData, setRowData] = React.useState<any[] | null>([]);
    const [searchText, setSearchText] = React.useState<string>("");
    const [loading, setLoading] = React.useState<boolean>(true);
    const [editingData, setEditingData] = React.useState<any>(null);
    const gridRef = React.useRef<AgGridReact<any>>(null);
    const [currentPage, setCurrentPage] = React.useState(1);
    const [totalPages, setTotalPages] = React.useState(1);
    const [isSearching, setIsSearching] = React.useState(false);
    const [isMarkDoneOpen, setIsMarkDoneOpen] = React.useState(false);
    const [debugSelectedRow, setDebugSelectedRow] = React.useState<any>(null);

    const gridComponents = React.useMemo<any>(() => {
        return {
            agColumnHeader: (props: any) => (
                <GridSortableHeader {...props}>
                    <TooltipTextCell value={props.displayName} />
                </GridSortableHeader>
            ),
            lookupComponent: LookupCellEditor
        };
    }, []);

    const gridAutoSizeStrategy = React.useMemo<any>(() => {
        return {
            type: "fitGridWidth",
            columnLimits: colDefs?.map(cd => ({
                colId: cd.field,
                minWidth: cd.width
            }))
        };
    }, [colDefs]);

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

    /**
     * Adds an "Actions" column to the grid for row-level actions.
     *
     * @param {ColDef[]} gridColumns - The existing grid columns.
     * @param {GridRowActionProps[]} rowActions - The actions available for each row.
     */
    function addActionColumn(gridColumns: (ColDef<any, any> & ColumnDefinition)[], rowActions: GridRowActionProps[]) {
        // Actions are always a third column
        gridColumns.splice(2, 0, {
            field: "actions",
            sortable: false,
            headerName: "Actions",
            headerClass: "eg-text-center",
            cellStyle: { textAlign: "center", pointerEvents: "auto !important" },
            minWidth: 80,
            maxWidth: 80,
            lockPinned: true,
            suppressMovable: true,
            cellRenderer: (p: CustomCellRendererProps) =>
                GridRowActions({
                    ...p,
                    actions: rowActions,
                    refreshGridData: refreshGridData,
                    entityName: props.entityName,
                    poHeaderId: props.headerId,
                    poHeaderSchemaName: props.headerSchemaName,
                    poHeaderEntityName: props.headerEntityName,
                    lineStatus: props.lineStatusLogicalName
                        ? gridColumns
                              .find(cd => cd.field === props.lineStatusLogicalName)
                              ?.options?.find(
                                  opt =>
                                      opt.label.toLowerCase() ===
                                      p.data[props.lineStatusLogicalName.toLowerCase()].toLowerCase()
                              )?.value ?? 0
                        : 0,
                    getString: props.getString,
                    attributesMetadata: props.attributesMetadata,
                    controlId: props.controlId,
                    webAPI: props.webApi,
                    debugMode: props.debugMode
                })
        });
    }

    /**
     * Determines the row style based on specific conditions.
     *
     * @param {RowClassParams} params - The row class parameters.
     * @returns {Object} The style object for the row.
     */
    const getRowStyle = React.useCallback((params: RowClassParams) => {
        const style: any = {};

        if (params.node.rowIndex === params.api.getDisplayedRowCount() - 1) {
            style.borderBottom = "solid 1px color-mix(in srgb, transparent, #181d1f 15%)";
        }

        if (params.data.statecode === "Inactive") {
            style.background = "color-mix(in srgb, transparent, lightgray 8%)";
        }

        return style;
    }, []);

    /**
     * Determines the row class based on specific conditions.
     *
     * @param {RowClassParams} params - The row class parameters.
     * @returns {string | undefined} The CSS class for the row.
     */
    const getRowClass = React.useCallback(
        (params: RowClassParams) => {
            if (props.pinIndex >= 0) {
                return "eg-pinned-selected";
            }
        },
        [props.pinIndex]
    );

    const getGridColumnsSetup = React.useCallback(() => {
        if (!retrievedColDefs) {
            return null;
        }

        const customizedColumns = customizeColumns(retrievedColDefs, props.customColumns);
        const gridColumns = getGridColumns(
            customizedColumns,
            retrievedRowData,
            props.readonly,
            props.lineStatusLogicalName,
            props.attributesMetadata,
            props.getString,
            props.webApi,
            props.hidePriceField,
            props.hidePriceFields,
            props.vendorId,
            props.headerEntityName === LineType.VEPendingVendorInvoiceHeader && !props.allowCreate ? true : false
        );

        if (!props.readonly && props.showRowActions) {
            addActionColumn(gridColumns, props.rowActions);
        }

        addDetailsColumn(gridColumns);
        addCommentsAndAttachmentsColumns(gridColumns);
        pinColumns(props.pinIndex, gridColumns);
        resizeColumns(gridRef.current?.api, gridColumns);

        return gridColumns;
    }, [
        retrievedColDefs,
        retrievedRowData,
        props.readonly,
        props.customColumns,
        props.showRowActions,
        props.rowActions
    ]);

    const addDetailsColumn = (gridColumns: (ColDef<any, any> & ColumnDefinition)[]) => {
        const fieldsToShow = gridColumns.filter(gc => gc.additionalDetails === true);

        if (!fieldsToShow.length || !retrievedRowData || !retrievedRowData.length) return;

        const fields = fieldsToShow.map(field => ({
            fieldName: field.field!,
            label:
                props.attributesMetadata.find(m => m.LogicalName === field.field)?.DisplayName?.UserLocalizedLabel
                    ?.Label ?? field.field
        }));

        gridColumns.push({
            field: "details",
            headerName: props.getString("DetailsGridColumn"),
            editable: false,
            sortable: false,
            cellRenderer: (params: any) => (
                <DetailsPopover
                    data={params.data}
                    additionalDetailsText={props.getString("AdditionalDetails")}
                    fields={fields}
                />
            ),
            minWidth: 80,
            width: 80
        });
    };

    const addCommentsAndAttachmentsColumns = (gridColumns: (ColDef<any, any> & ColumnDefinition)[]) => {
        if (!props.showCommentsAndAttachments || !retrievedRowData || !retrievedRowData.length) return;

        const sampleRow = retrievedRowData[0];

        if (props.commentsCountField && sampleRow.hasOwnProperty(props.commentsCountField)) {
            gridColumns.push({
                field: props.commentsCountField,
                headerName: props.getString("NoOfComments"),
                editable: false,
                sortable: false,
                cellRenderer: (params: any) => (
                    <span style={{ display: "flex", alignItems: "center", justifyContent: "cetner", gap: "4px" }}>
                        <CommentMultipleRegular />
                        {params.data?.[props.commentsCountField] ?? 0}
                    </span>
                ),
                minWidth: 112,
                width: 112
            });
        }

        if (props.attachmentsCountField && sampleRow.hasOwnProperty(props.attachmentsCountField)) {
            gridColumns.push({
                field: props.attachmentsCountField,
                headerName: props.getString("NoOfAttachments"),
                editable: false,
                sortable: false,
                cellRenderer: (params: any) => (
                    <span style={{ display: "flex", alignItems: "center", justifyContent: "cetner", gap: "4px" }}>
                        <AttachRegular />
                        {params.data?.[props.attachmentsCountField] ?? 0}
                    </span>
                ),
                minWidth: 112,
                width: 112
            });
        }
    };

    /**
     * Handles the initialization of the grid and retrieves column and row data.
     */
    const handle_Grid_Ready = React.useCallback(async () => {
        const retrievedColDefs = await props.retrieveColumnDefs(
            props.customColumns.filter(c => c.dataType === "lookup") || []
        );
        const retrievedRows = await props.retrieveRows(false, 1);

        for (const row of retrievedRows.data) {
            convertRowData.dataverseToGridApp(row, retrievedColDefs);
        }

        setRetrievedColDefs(retrievedColDefs);
        setRetrievedRowData(retrievedRows.data.map(r => ({ ...r })));
        setRowData(retrievedRows.data);
        setCurrentPage(1);
        setTotalPages(Math.ceil(retrievedRows.totalRecordCount / 10));
        setLoading(false);
    }, []);

    /**
     * Refreshes the grid data.
     *
     * This is called when a row action has been executed
     */
    const refreshGridData = React.useCallback(async () => {
        setLoading(true);
        const retrievedColDefs = await props.retrieveColumnDefs(
            props.customColumns.filter(c => c.dataType === "lookup") || []
        );
        const retrievedRows = await props.retrieveRows(isSearching, currentPage, searchText);

        for (const row of retrievedRows.data) {
            convertRowData.dataverseToGridApp(row, retrievedColDefs);
        }

        setRetrievedColDefs(retrievedColDefs);
        setRetrievedRowData(retrievedRows.data.map(r => ({ ...r })));
        setRowData(retrievedRows.data);
        setTotalPages(Math.ceil(retrievedRows.totalRecordCount / 10));
        setLoading(false);
    }, [currentPage, isSearching, searchText]);

    /**
     * Handles column visibility toggling when a user interacts with a column toggle switch.
     *
     * @param {string} field - The column field.
     * @param {SwitchOnChangeData} data - The switch event data.
     */
    const handle_ColumnSwitch_Change = React.useCallback(
        (field: string, data: SwitchOnChangeData) => {
            if (!colDefs) {
                return;
            }

            const index = colDefs?.findIndex(cd => cd.field === field);

            const newCD = { ...colDefs[index] };

            newCD.hide = !data.checked;

            const newColDefs = [...colDefs.slice(0, index), newCD, ...colDefs.splice(index + 1)];

            setColDefs(newColDefs);
        },
        [colDefs]
    );

    /**
     * Handles search functionality for filtering row data.
     *
     * @param {SearchBoxChangeEvent} event - The search event.
     * @param {InputOnChangeData} data - The input event data.
     */
    const handle_SearchBox_Change = async (event: SearchBoxChangeEvent | React.MouseEvent, data: InputOnChangeData) => {
        setSearchText(data.value);
        if (event.type === "click" && data.value === "") {
            setLoading(true);
            const retrievedColDefs = await props.retrieveColumnDefs(
                props.customColumns.filter(c => c.dataType === "lookup") || []
            );
            const retrievedRows = await props.retrieveRows(false, 1);

            for (const row of retrievedRows.data) {
                convertRowData.dataverseToGridApp(row, retrievedColDefs);
            }

            setRetrievedColDefs(retrievedColDefs);
            setRetrievedRowData(retrievedRows.data.map(r => ({ ...r })));
            setRowData(retrievedRows.data);
            setCurrentPage(1);
            setTotalPages(Math.ceil(retrievedRows.totalRecordCount / 10));
            setLoading(false);
        }
    };

    /**
     * Handles searching when enter key has been pressed on the search box
     */
    const handleEnterKeyPress = React.useCallback(
        async (event: React.KeyboardEvent<HTMLInputElement>) => {
            if (event.key === "Enter") {
                if (event.target instanceof HTMLInputElement && event.target.type === "search") event.preventDefault();

                setLoading(true);
                setIsSearching(true);
                const retrievedColDefs = await props.retrieveColumnDefs(
                    props.customColumns.filter(c => c.dataType === "lookup") || []
                );
                const retrievedRows = await props.retrieveRows(true, 1, searchText);

                for (const row of retrievedRows.data) {
                    convertRowData.dataverseToGridApp(row, retrievedColDefs);
                }

                setRetrievedColDefs(retrievedColDefs);
                setRetrievedRowData(retrievedRows.data.map(r => ({ ...r })));
                setRowData(retrievedRows.data);
                setCurrentPage(1);
                setTotalPages(Math.ceil(retrievedRows.totalRecordCount / 10));
                setLoading(false);
            }
        },
        [searchText, isSearching]
    );

    /**
     * Makes a deep clone object from the data that has been changed.
     *
     * Only columns that are allowed to updated are stored in this object.
     * This object is then compared to the new values on the RowEditingStopped event
     * @param event The row event
     */
    const handleRowDataEditingStarted = (event: RowEditingStartedEvent<any>) => {
        const updateColumns = colDefs.filter(cd => cd.allowUpdate === true).map(cd => cd.field);
        setEditingData(
            cloneDeep(
                updateColumns.reduce((acc, field) => {
                    const cd = colDefs.find(c => c.field === field);
                    if (event.data.hasOwnProperty(field)) {
                        if (event.data[field] && cd.dataType === "picklist" && cd.options && !cd.showAsCheckbox) {
                            acc[field] = cd.options?.find(opt => opt.label === event.data[field]).value;
                        } else if (
                            event.data[field] &&
                            cd.dataType === "datetime" &&
                            event.data[field] instanceof Date
                        ) {
                            acc[field] = new Date(event.data[field].getTime());
                        } else if (event.data[field] && cd.dataType === "integer") {
                            acc[field] = parseInt(event.data[field]);
                        } else if (
                            (event.data[field] && cd.dataType === "decimal") ||
                            cd.dataType === "double" ||
                            cd.dataType === "money"
                        ) {
                            acc[field] = parseFloat(event.data[field]);
                        } else acc[field] = event.data[field];
                    }
                    return acc;
                }, {} as any)
            )
        );
    };

    /**
     * Handles row value changes and updates data accordingly.
     *
     * @param {RowEditingStoppedEvent<any, any>} event - The row change event.
     */
    const handleRowDataEditingStopped = async (event: RowEditingStoppedEvent<any>) => {
        if (event.event instanceof KeyboardEvent && event.event.key.toLowerCase() === "escape") {
            event.node.setSelected(false);
            return;
        }
        setLoading(true);
        if (!editingData) {
            setLoading(false);
            return;
        }

        const updateColumns = colDefs.filter(cd => cd.allowUpdate === true || cd.editableOnAlternate === true);

        // Compare values and store only changed fields
        const updateData = updateColumns.reduce((acc, col) => {
            const field = col.field;
            const isPicklist = col.dataType === "picklist";

            // Get the original and new values
            const oldValue = isPicklist
                ? col.showAsCheckbox
                    ? editingData[field]
                    : col.options?.find(opt => opt.value === editingData[field])?.value
                : col.dataType === "lookup"
                ? editingData[field]?.id ?? null
                : editingData[field];

            const newValue = isPicklist
                ? col.showAsCheckbox
                    ? event.data[field]
                    : col.options?.find(opt => opt.label === event.data[field])?.value
                : col.dataType === "lookup"
                ? event.data[field]?.id ?? null
                : event.data[field];

            let valueChanged = false;

            if (oldValue instanceof Date && newValue instanceof Date) {
                valueChanged = oldValue.getTime() !== newValue.getTime();
            } else if (
                (typeof oldValue === "number" || typeof newValue === "number") &&
                parseFloat(String(oldValue)) === parseFloat(String(newValue))
            ) {
                valueChanged = false;
            } else {
                valueChanged = oldValue !== newValue;
            }

            if (valueChanged) {
                if (newValue instanceof Date) {
                    acc[field] = `${newValue.getFullYear()}-${(newValue.getMonth() + 1)
                        .toString()
                        .padStart(2, "0")}-${newValue.getDate().toString().padStart(2, "0")}`;
                } else if (col.dataType === "lookup") {
                    if (newValue)
                        acc[
                            `${props.attributesMetadata.find(m => m.LogicalName === field)?.SchemaName}@odata.bind`
                        ] = `/${_getPlural(
                            props.attributesMetadata.find(m => m.LogicalName === field)?.Targets?.[0]
                        )}(${newValue})`;
                    else acc[field] = newValue;
                } else acc[field] = newValue;
            }

            return acc;
        }, {} as any);

        const idKey = props.entityName + "id";
        const recordId = event.data[idKey];

        if (Object.keys(updateData).length === 0) {
            setLoading(false);
            return;
        }

        if (!updateData.hasOwnProperty(`${props.headerSchemaName}@odata.bind`) && event.data.isCreate) {
            updateData[`${props.headerSchemaName}@odata.bind`] = `/${_getPlural(props.headerEntityName)}(${
                props.headerId
            })`;
        }

        await props.onRowValueChanged(recordId, updateData, event.data.isCreate ?? false, setLoading);

        const retrievedColDefs = await props.retrieveColumnDefs(
            props.customColumns.filter(c => c.dataType === "lookup") || []
        );
        const retrievedRows = await props.retrieveRows(isSearching, currentPage, searchText);

        for (const row of retrievedRows.data) {
            convertRowData.dataverseToGridApp(row, retrievedColDefs);
        }

        setRetrievedColDefs(retrievedColDefs);
        setRetrievedRowData(retrievedRows.data.map(r => ({ ...r })));
        setRowData(retrievedRows.data);
        setTotalPages(Math.ceil(retrievedRows.totalRecordCount / 10));
        setLoading(false);
    };

    /**
     * Adds a new row into the AG Grid
     */
    const addNewRow = () => {
        const newItem: any = {};

        for (const col of colDefs) {
            const attr = props.entityMetadata?.Attributes.get(col.field);
            const isWritable = attr?.IsValidForCreate ?? true; // assume true if metadata missing

            if (!isWritable) continue;

            switch (col.dataType) {
                case "string":
                    newItem[col.field] = "";
                    break;
                case "integer":
                case "decimal":
                case "double":
                case "money":
                    newItem[col.field] = 0;
                    break;
                case "boolean":
                    newItem[col.field] = false;
                    break;
                case "datetime":
                    newItem[col.field] = new Date();
                    break;
                case "picklist":
                    newItem[col.field] = col.options?.[0]?.label ?? "";
                    break;
                default:
                    newItem[col.field] = null;
            }
        }

        newItem.isCreate = true;
        newItem.editable = true;
        newItem.allowUpdate = true;
        newItem[`${props.headerSchemaName}@odata.bind`] = `/${_getPlural(props.headerEntityName)}(${props.headerId})`;

        gridRef.current!.api.applyTransaction({ add: [newItem] });
        setRowData(prev => [...prev, newItem]);
        setRetrievedRowData(prev => [...(prev ?? []), { ...newItem }]);

        let lastRowIndex = gridRef.current!.api.getDisplayedRowCount() - 1;
        if (lastRowIndex < 0) {
            lastRowIndex = rowData.length - 1;
        }
        if (lastRowIndex < 0) lastRowIndex = 0;

        const firstEditableCol = colDefs.find(c => c.field && newItem.hasOwnProperty(c.field))?.field;
        if (firstEditableCol && gridRef.current) {
            gridRef.current!.api.ensureIndexVisible(lastRowIndex);
            gridRef.current!.api.setFocusedCell(lastRowIndex, firstEditableCol);
            gridRef.current!.api.startEditingCell({
                rowIndex: lastRowIndex,
                colKey: firstEditableCol
            });
        }
    };

    /**
     * Dispatches a custom event "vrm:pcfCommunication"
     *
     * This event is used in the Attachments PCF control and switches
     * the tab to lines, retrieving the attachments for the Line record
     *
     * @param row The row event
     */
    const handleRowClick = (row: RowSelectedEvent<any>) => {
        if (row.node.isSelected()) {
            let alternateLineId = null;
            if (props.headerEntityName === LineType.VERFQReplyHeader) {
                if (row.data.hasOwnProperty(props.alternateLineFieldName)) {
                    alternateLineId = row.data?.[props.alternateLineFieldName]?.id
                        ? row.data[props.alternateLineFieldName].id
                        : null;
                } else if (row.data.hasOwnProperty(`_${props.alternateLineFieldName}_value`)) {
                    alternateLineId = row.data[`_${props.alternateLineFieldName}_value`];
                }
            }

            var customEvent = new CustomEvent<PCFCommunication>(
                `${PCFCommunicationType.Communication}:${props.controlId}`,
                {
                    detail: {
                        id: row.data[props.entityName + "id"],
                        lineNumberFieldName: AttachmentFieldName.LineNumber[props.entityName],
                        name:
                            typeof row.data[AttachmentFieldName.LineNumber[props.entityName]] === "string"
                                ? parseFloat(row.data[AttachmentFieldName.LineNumber[props.entityName]])
                                : row.data.hasOwnProperty(AttachmentFieldName.LineNumber[props.entityName])
                                ? typeof row.data[AttachmentFieldName.LineNumber[props.entityName]] === "string"
                                    ? parseFloat(row.data[AttachmentFieldName.LineNumber[props.entityName]])
                                    : row.data[AttachmentFieldName.LineNumber[props.entityName]]
                                : row.data.hasOwnProperty("mserp_linenumber")
                                ? typeof row.data.mserp_linenumber === "string"
                                    ? parseFloat(row.data.mserp_linenumber)
                                    : row.data.mserp_linenumber
                                : 0,
                        logicalName: props.entityName,
                        attachmentEntityName: AttachmentEntityName[props.entityName],
                        attachmentLineLookupName: AttachmentFieldName.Lookup[props.entityName],
                        alternateLineId: alternateLineId
                    }
                }
            );
            window.dispatchEvent(customEvent);
        }
        setDebugSelectedRow(row.data);
    };

    /**
     * Handles simple pagination
     * @param newPage The new page number
     */
    const handlePageChange = async (newPage: number) => {
        setCurrentPage(newPage);
        setLoading(true);

        const retrievedColDefs = await props.retrieveColumnDefs(
            props.customColumns.filter(c => c.dataType === "lookup") || []
        );
        const retrievedRows = await props.retrieveRows(isSearching, newPage, searchText);

        for (const row of retrievedRows.data) {
            convertRowData.dataverseToGridApp(row, retrievedColDefs);
        }

        setRetrievedColDefs(retrievedColDefs);
        setRetrievedRowData(retrievedRows.data.map(r => ({ ...r })));
        setRowData(retrievedRows.data);
        setTotalPages(Math.ceil(retrievedRows.totalRecordCount / 10));
        setLoading(false);
    };

    /**
     * Calls OData action in F&O to mark all line records as Done.
     */
    const handleMarkAllAsDone = async () => {
        try {
            await executeAction(
                VirtualEntityActions[props.entityName].MarkAsDone,
                props.headerEntityName,
                props.headerId,
                VirtualEntityActions[props.entityName].MarkAsDone,
                props.webApi
            );

            refreshGridData();
        } catch (e) {
            alert(e);
        }
    };

    React.useEffect(() => {
        const gridColumns = getGridColumnsSetup();
        if (gridColumns) setColDefs(gridColumns);
    }, [getGridColumnsSetup]);

    React.useEffect(() => {
        resizeColumns(gridRef.current?.api, colDefs);
    }, [colDefs]);

    React.useEffect(() => {
        const debouncedHandleResize = debounce(() => resizeColumns(gridRef.current?.api, colDefs), 100);
        const handleCustomControlRefreshEvent = async (event: CustomEvent): Promise<void> => {
            if (event && event.detail && event.detail?.refresh) {
                await refreshGridData();
            }
        };

        /**
         * Handles clicking outside of the Grid and anywhere on the page.
         * If the row is in edit, it will stop editing and save the record
         * @param event Mouse event
         */
        const handleClickOutside = (event: MouseEvent) => {
            const path = event.composedPath?.() || [];

            const clickedInsideControl = path.includes(props.controlRef.current);
            const clickedInsideLookup = path.some(
                el =>
                    el instanceof HTMLElement &&
                    (el.classList.contains("custom-lookup-popover") ||
                        el.classList.contains("fui-PopoverSurface") ||
                        el.hasAttribute("data-lookup-popup"))
            );

            if (!clickedInsideControl && !clickedInsideLookup) {
                const isEditing = gridRef!.current?.api.getEditingCells()?.length;
                if (isEditing) {
                    gridRef!.current?.api.stopEditing();
                }
            }

            const selectedNodes = gridRef!.current?.api.getSelectedNodes();
            if (selectedNodes && selectedNodes.length > 0) {
                gridRef!.current?.api.deselectAll();
            }
        };

        window.addEventListener("resize", debouncedHandleResize);
        window.addEventListener(
            `${PCFCommunicationType.RefreshLines}:${props.controlId}`,
            handleCustomControlRefreshEvent
        );
        document.addEventListener("mousedown", handleClickOutside);

        return () => {
            window.removeEventListener("resize", debouncedHandleResize);
            window.removeEventListener(
                `${PCFCommunicationType.RefreshLines}:${props.controlId}`,
                handleCustomControlRefreshEvent
            );
            document.removeEventListener("mousedown", handleClickOutside);
        };
    });

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

    return (
        <FluentProvider theme={webLightTheme}>
            <div className="eg-toolbar">
                <SearchBox
                    className="eg-searchbox"
                    placeholder={props.getString("Search")}
                    appearance="outline"
                    value={searchText}
                    onChange={handle_SearchBox_Change}
                    onKeyDown={handleEnterKeyPress}
                />

                {props.showMarkAllButton && (
                    props.headerStatus === FieldNames.HeaderStatusEdited[LineType.VEPurchaseOrderResponseHeader] ||
                    props.headerStatus === FieldNames.HeaderStatusEdited[LineType.VERFQReplyHeader]
                ) && (
                    <Button onClick={() => setIsMarkDoneOpen(true)}>{props.getString("MarkAllAsDone")}</Button>
                )}

                {props.headerEntityName === LineType.VEPendingVendorInvoiceHeader && props.allowCreate && (
                    <div className="row-buttons">
                        <Button icon={<AddRegular />} onClick={addNewRow}>
                            {props.getString("NewInvoice")}
                        </Button>
                    </div>
                )}

                {colDefs
                    ?.filter(cd => cd.toggleShowColumn)
                    .map(cd => (
                        <Switch
                            className="eg-column-toggle"
                            label={cd.headerName}
                            key={`eg-column-toggle-${cd.field}`}
                            defaultChecked={cd.toggleShowColumnEnabled !== false}
                            onChange={(e, d) => handle_ColumnSwitch_Change(cd.field!, d)}
                        />
                    ))}
            </div>

            <div className="eg-grid">
                <AgGridReact
                    ref={gridRef}
                    theme={gridTheme}
                    loading={loading}
                    loadingOverlayComponent={CustomLoadingOverlay}
                    rowData={rowData}
                    columnDefs={colDefs}
                    components={gridComponents}
                    autoSizeStrategy={gridAutoSizeStrategy}
                    processUnpinnedColumns={() => []}
                    tooltipShowDelay={500}
                    tooltipHideDelay={1500}
                    singleClickEdit={true}
                    suppressPaginationPanel
                    paginationPageSizeSelector={false}
                    rowSelection="single"
                    editType="fullRow"
                    getRowStyle={getRowStyle}
                    getRowClass={getRowClass}
                    onGridReady={handle_Grid_Ready}
                    onRowEditingStarted={handleRowDataEditingStarted}
                    onRowEditingStopped={handleRowDataEditingStopped}
                    onRowSelected={handleRowClick}
                    onCellClicked={handleCellClick}
                    domLayout="autoHeight"
                />
            </div>
            <div>
                <CustomPagination
                    currentPage={currentPage}
                    totalPages={totalPages}
                    onPageChange={handlePageChange}
                    getString={props.getString}
                />
            </div>

            <ConfirmationDialog
                isOpen={isMarkDoneOpen}
                title={props.getString("MarkAllAsDoneConfirmation")}
                content={""}
                onConfirm={handleMarkAllAsDone}
                onClose={() => setIsMarkDoneOpen(false)}
                getString={props.getString}
                confirmText={props.getString("Yes")}
                cancelText={props.getString("No")}
            />

            {props.debugMode && <div className="eg-debug-panel">
                <div className="eg-debug-panel__header">DEBUG v1.2 (click any row)</div>
                {(() => {
                    const portalLang = (window as any)?.msdyn?.Portal?.Snippets?.userLanguage ||
                        (parent?.window as any)?.msdyn?.Portal?.Snippets?.userLanguage;
                    const navLocale = navigator.language;
                    const locale = portalLang?.bcp47_locale || navLocale || "en-US";
                    const groupingSep = (portalLang?.thousand_separator ?? new Intl.NumberFormat(locale, { useGrouping: true }).format(1111)[1]).trim();
                    const decimalSep = (portalLang?.decimal_separator ?? new Intl.NumberFormat(locale, { minimumFractionDigits: 1 }).format(1.1)[1]).trim();
                    const numericTypes = ["decimal", "double", "money", "integer"];
                    const numericCols = (retrievedColDefs || []).filter(c => numericTypes.includes(c.dataType));
                    const isRawApiValue = (v: any) => typeof v === "string" && /^\d+\.\d+$/.test(v);
                    return (
                        <div className="eg-debug-panel__body">
                            <div className="eg-debug-panel__locale">
                                <b>Locale:</b> portal=<code>{portalLang?.bcp47_locale ?? "(none)"}</code>{" "}
                                nav=<code>{navLocale}</code>{" "}
                                thousands=<code>"{groupingSep}"</code>{" "}
                                decimal=<code>"{decimalSep}"</code>
                            </div>
                            {debugSelectedRow ? (
                                <table className="eg-debug-panel__table">
                                    <thead>
                                        <tr>
                                            <th className="eg-debug-panel__th">Field</th>
                                            <th className="eg-debug-panel__th">row.data (type)</th>
                                            <th className="eg-debug-panel__th">raw API?</th>
                                            <th className="eg-debug-panel__th">parseFloat</th>
                                            <th className="eg-debug-panel__th">parseDecimal</th>
                                            <th className="eg-debug-panel__th">valueGetter</th>
                                            <th className="eg-debug-panel__th">formatted</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        {numericCols.map(col => {
                                            const raw = debugSelectedRow[col.field];
                                            const rawApi = isRawApiValue(raw);
                                            const vgResult = typeof raw === "string"
                                                ? (col.convertToInteger ? parseFloat(raw).toFixed(0) : parseFloat(raw).toFixed(2))
                                                : typeof raw === "number"
                                                ? (col.convertToInteger ? raw.toFixed(0) : raw.toFixed(2))
                                                : String(raw ?? "");
                                            const vfNum = Number(vgResult);
                                            const vfResult = isNaN(vfNum) ? "" : col.dataType === "integer"
                                                ? formatInteger(vfNum)
                                                : formatDecimal(vfNum, col.convertToInteger ? 0 : 2);
                                            return (
                                                <tr key={col.field}>
                                                    <td className="eg-debug-panel__td">{col.field}</td>
                                                    <td className="eg-debug-panel__td"><code>"{String(raw)}" ({typeof raw})</code></td>
                                                    <td className="eg-debug-panel__td"><code>{rawApi ? "YES" : "NO"}</code></td>
                                                    <td className="eg-debug-panel__td"><code>{parseFloat(String(raw))}</code></td>
                                                    <td className="eg-debug-panel__td"><code>{parseDecimal(raw)}</code></td>
                                                    <td className="eg-debug-panel__td"><code>"{vgResult}"</code></td>
                                                    <td className="eg-debug-panel__td"><code>"{vfResult}"</code></td>
                                                </tr>
                                            );
                                        })}
                                    </tbody>
                                </table>
                            ) : (
                                <div className="eg-debug-panel__empty">Click any row to inspect its numeric field values.</div>
                            )}
                        </div>
                    );
                })()}
            </div>}
        </FluentProvider>
    );
};
