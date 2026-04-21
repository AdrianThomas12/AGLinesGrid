/**
 * @file renderer.ts
 * @description This file handles the rendering process of the application, managing interactions and updates within the renderer process.
 */

import * as ReactDOM from "react-dom";
import * as React from "react";
import { App } from "../Components/App";
import * as POConfiguration from "./configuration";
import * as GridWebApi from "./webApi";
import { IInputs } from "../generated/ManifestTypes";
import {
    BooleanOptions,
    ColumnDefinition,
    FieldNames,
    GridProps,
    LineStatusOptions,
    LineType,
    StateOptions
} from "./types";
import {
    buildLookupEntityColumns,
    getEntityDto,
    getEntityMetadataUnified,
    getLookupTargetsFromPortalDefs
} from "./webApi";

/**
 * Class responsible for rendering and managing the editable grid component.
 */
export class EditableGridRenderer {
    private attributeList: string[] = [];

    constructor() { }

    /**
     * Renders the application server-side with provided configuration.
     * @param {HTMLElement} container - The container element where the app will be rendered.
     * @param {IInputs} context - The context data for rendering.
     */
    public async renderServer(
        context: ComponentFramework.Context<IInputs>,
        container: HTMLDivElement,
        headerId: string,
        headerEntityName: string,
        config: GridProps
    ) {
        const pageNumber = 1;
        const poHeaderSchemaName = context.parameters.headerSchemaName.raw
            ? context.parameters.headerSchemaName.raw
            : "mserp_FK_VRMPurchaseOrderResponseHeader_id";
        const { fetchxml, layoutjson } = await this.getViewData(context, poHeaderSchemaName, headerId);
        const gridEntityName = this.getEntityNameFromFetchXml(fetchxml);
        const gridColumns = layoutjson.Rows[0].Cells.map((c: any) => c.Name);

        const entityMetadata = await getEntityMetadataUnified(
            context,
            gridEntityName,
            [...gridColumns],
            context.userSettings.languageId ?? 1033
        );
        const attributesMetadata = entityMetadata.Attributes.get();

        const edited: number = context.parameters.status_edited.raw
            ? context.parameters.status_edited.raw
            : FieldNames.HeaderStatusEdited[headerEntityName];
        const headerStatusFieldName =
            context.parameters.status.attributes?.LogicalName?.toLowerCase() ??
            FieldNames.HeaderStatus[headerEntityName];
        const headerStatus: number =
            context.parameters.status.raw ??
            parseInt((document.getElementById(headerStatusFieldName) as HTMLSelectElement)?.value ?? "0");
        const readOnly = headerStatus !== edited;
        let allowCreate = false;
        let vendorId = null;

        if (headerEntityName === LineType.VEPendingVendorInvoiceHeader) {
            allowCreate = await GridWebApi.checkInvoicesPurchaseOrder(
                headerEntityName,
                headerId,
                "mserp_fk_purchaseorderheader_id",
                context.webAPI
            );

            vendorId = await GridWebApi.getVendorForInvoice(
                headerEntityName,
                headerId,
                "mserp_fk_vendor_id",
                context.webAPI
            );
        }

        const retrieveColumnDefs = async (lookups: ColumnDefinition[]) => {
            return await this.getGridColumnDefs(
                context,
                gridEntityName,
                gridColumns,
                layoutjson,
                attributesMetadata,
                lookups
            );
        };

        const retrieveRows = async (isSearching: boolean, page: number, searchText?: string) => {
            const searchableColumns = context.parameters.searchableColumns.raw
                ? context.parameters.searchableColumns.raw.split(",").map(s => s.trim())
                : [];

            if (isSearching) {
                const filteredFetchXml = this.addFiltersOnSearch(
                    searchText,
                    attributesMetadata,
                    searchableColumns,
                    fetchxml
                );
                return await this.getGridRows(gridEntityName, gridColumns, filteredFetchXml, page, context.webAPI);
            } else return await this.getGridRows(gridEntityName, gridColumns, fetchxml, page, context.webAPI);
        };

        const handleRowValueChanged = async (
            id: string,
            data: any,
            isCreate: boolean,
            setLoading: React.Dispatch<React.SetStateAction<boolean>>
        ) => {
            try {
                if (isCreate) {
                    await GridWebApi.createRecord(gridEntityName, data, context.webAPI, headerEntityName, headerId);
                } else {
                    await GridWebApi.updateRecord(gridEntityName, id, data, context.webAPI);
                }
            } catch (error) {
                setLoading(false);
                alert(error);
            }
        };

        const handleDeleteRow = async (id: string) => {
            try {
                await GridWebApi.deleteRecord(gridEntityName, id, context.webAPI);
            } catch (error) {
                alert(error);
            }
        };

        ReactDOM.render(
            React.createElement(App, {
                entityName: gridEntityName,
                headerId: headerId,
                headerSchemaName: poHeaderSchemaName,
                headerEntityName: headerEntityName,
                readonly: context.mode.isControlDisabled || readOnly,
                showRowActions: config.showRowActions ?? true,
                rowActions: POConfiguration.getRowActions(gridEntityName, {
                    handleDeleteRow,
                    executeServerAction: (gridRowActionName: string, id: string, actionName: string, data?: any) =>
                        GridWebApi.executeAction(
                            gridRowActionName,
                            gridEntityName,
                            id,
                            actionName,
                            context.webAPI,
                            data
                        ),
                    getString: (str: string) => context.resources.getString(str)
                }),
                fitToContainer: false,
                retrieveColumnDefs: retrieveColumnDefs,
                retrieveRows: retrieveRows,
                onRowValueChanged: handleRowValueChanged,
                pageNumber: pageNumber,
                lineStatusLogicalName: context.parameters.recordStatusLogicalName.raw
                    ? context.parameters.recordStatusLogicalName.raw
                    : "mserp_linestate",
                attributesMetadata: attributesMetadata,
                getString: (s: string) => context.resources.getString(s),
                controlId: context.parameters.controlId.raw,
                webApi: context.webAPI,
                alternateLineFieldName: context.parameters.alternateLineId.raw,
                allowCreate: allowCreate,
                vendorId: vendorId,
                headerStatus: headerStatus,
                debugMode: context.parameters.debugMode?.raw ?? false,
                ...config
            }),
            container
        );
    }

    /**
     * Retrieves the data for the grid view.
     * @returns {Promise<any>} A promise resolving with the grid view data.
     */
    private async getViewData(
        context: ComponentFramework.Context<IInputs>,
        headerLogicalName: string,
        entityId: string
    ): Promise<any> {
        const viewId = context.parameters.linesDataset.getViewId();

        const view = await GridWebApi.retrieve("savedquery", viewId, context.webAPI);
        const fetchxml: string = this.addPOHeaderFilterFetchXml(
            headerLogicalName,
            entityId,
            view.fetchxml,
            context.parameters.clearCacheProperty.raw
        );
        const layoutjson = JSON.parse(view.layoutjson);

        return { fetchxml, layoutjson };
    }

    /**
     * Retrieves column definitions for the grid.
     * @returns {Promise<ColumnDefinition[]>} A promise resolving with the grid column definitions.
     */
    private async getGridColumnDefs(
        context: ComponentFramework.Context<IInputs>,
        entityName: string,
        gridColumns: string[],
        layoutjson: any,
        attributesMetadata: any[],
        lookups: ColumnDefinition[]
    ) {
        const columnDefs: ColumnDefinition[] = [];

        for (const c of layoutjson.Rows[0].Cells) {
            const attrName = c.Name;
            const attrMetadata = attributesMetadata.find((m: any) => m.LogicalName === c.Name);

            const header = attrMetadata.DisplayName?.UserLocalizedLabel?.Label
                ? attrMetadata.DisplayName.UserLocalizedLabel.Label
                : attrMetadata.DisplayName;
            const dataType =
                attrMetadata.AttributeType && typeof attrMetadata.AttributeType === "string"
                    ? attrMetadata.AttributeType?.toLowerCase()
                    : attrMetadata.AttributeTypeName?.toLowerCase();
            const options =
                dataType == "picklist" || dataType === "status"
                    ? undefined
                    : dataType === "state"
                        ? StateOptions
                        : dataType === "boolean"
                            ? BooleanOptions
                            : undefined;

            let columns = null;
            if (dataType.toLowerCase() === "lookup") {
                // 1) Resolve lookup targets from portal entity definitions
                const targets = getLookupTargetsFromPortalDefs(entityName, attrName);
                const lookupEntityType = targets[0];

                if (!lookupEntityType) {
                    throw new Error(`Lookup targets not found for ${entityName}.${attrName}`);
                }

                const lookupDef = lookups.find(l => l.field === attrName);

                // 2) Determine which columns to show
                const lookupColumns = lookupDef?.lookupColumns?.length
                    ? lookupDef.lookupColumns
                    : (() => {
                        // Try to get primary name from portal defs
                        const ent = getEntityDto(lookupEntityType);
                        return ent?.primaryNameAttribute ? [ent.primaryNameAttribute] : []; // fallback handled below
                    })();

                // 3) Build columns via portal bundle (with fallback inside helper)
                if (lookupColumns.length > 0) {
                    columns = await buildLookupEntityColumns(
                        context,
                        lookupEntityType,
                        lookupColumns,
                        context.userSettings?.languageId
                    );
                } else {
                    // Absolute fallback: Name
                    columns = [{ key: "name", label: "Name" }];
                }
            }

            columnDefs.push({
                field: attrName,
                headerName: header,
                dataType: dataType,
                width: c.Width,
                options: options,
                columns: columns
            });
        }

        return columnDefs;
    }

    /**
     * Retrieves all rows for the grid.
     * @returns {Promise<any[]>} A promise resolving with an array of grid rows.
     */
    private async getGridRows(
        entityName: string,
        gridColumns: string[],
        fetchxml: string,
        page: number,
        contextWebApi: ComponentFramework.WebApi
    ): Promise<{ data: any; totalRecordCount: number }> {
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(fetchxml, "text/xml");
        const fetchNode = xmlDoc.getElementsByTagName("fetch")[0];
        if (fetchNode) {
            fetchNode.setAttribute("returntotalrecordcount", "true");
            fetchNode.setAttribute("page", page.toString());
            fetchNode.setAttribute("count", "10");
            fetchNode.removeAttribute("savedqueryid");
        }

        const serializer = new XMLSerializer();
        let reqFetch = serializer.serializeToString(xmlDoc);

        if (entityName === "mserp_vrmpendingvendorinvoicelinedetailsentity") {
            reqFetch = this.addProcurementCategoryLinkedEntity(reqFetch);
        }

        const serverRecords = await GridWebApi.retrieveMultiple(
            entityName,
            "?fetchXml=" + encodeURIComponent(reqFetch),
            contextWebApi
        );
        const records = serverRecords.value || serverRecords.entities || [];

        const gridRows: any[] = [];
        let totalRecordCount = 0;

        for (const serverRec of records) {
            const row: any = this.getGridRow(serverRec, entityName, gridColumns);

            gridRows.push(row);
        }

        if (serverRecords.hasOwnProperty("@Microsoft.Dynamics.CRM.totalrecordcount")) {
            totalRecordCount = serverRecords["@Microsoft.Dynamics.CRM.totalrecordcount"];
        }

        return { data: gridRows, totalRecordCount: totalRecordCount };
    }

    /**
     * Injects procurement category metadata into the FetchXML for specific entity types.
     * This method adds the `<all-attributes />` tag to ensure all fields are retrieved 
     * and performs an outer join with the procurement product category assignment entity 
     * to retrieve the specific category name.
     * * @param {string} reqFetch - The original FetchXML string to be modified.
     * @returns {string} The modified FetchXML string containing the new nodes and link-entity.
     */
    private addProcurementCategoryLinkedEntity(reqFetch: string) {
        // 1. Parse the existing FetchXML string into an XML Document
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(reqFetch, "text/xml");
        const entityNode = xmlDoc.getElementsByTagName("entity")[0];

        if (entityNode) {
            // 1. Create and add the <all-attributes /> node
            const allAttributes = xmlDoc.createElement("all-attributes");
            entityNode.appendChild(allAttributes);
            // 2. Create the link-entity element
            const linkEntity = xmlDoc.createElement("link-entity");
            linkEntity.setAttribute("alias", "a");
            linkEntity.setAttribute("name", "mserp_vrmprocurementproductcategoryvendorassignmententity");
            linkEntity.setAttribute("to", "mserp_fk_vendorapprovedprocurementcategory_id");
            linkEntity.setAttribute("from", "mserp_vrmprocurementproductcategoryvendorassignmententityid");
            linkEntity.setAttribute("link-type", "outer");
            linkEntity.setAttribute("visible", "false");

            // 3. Create the attribute element inside the link-entity
            const attr = xmlDoc.createElement("attribute");
            attr.setAttribute("name", "mserp_productcategoryname");

            // 4. Assemble and append to the entity
            linkEntity.appendChild(attr);
            entityNode.appendChild(linkEntity);
        }

        // 5. Serialize back to string
        const finalFetchXml = new XMLSerializer().serializeToString(xmlDoc);

        return finalFetchXml;
    }

    /**
     * Transforms raw row data from the server into a format compatible with the grid application.
     * This method ensures that date fields are properly converted and formatted.
     *
     * @param {any} serverRecord - The raw row data object fetched from the server.
     * @param {string} entityName - The name of the entity associated with the data.
     * @param {string[]} gridColumns - The list of column names to be included in the grid.
     * @returns {any} The processed row data formatted for the grid component.
     */
    private getGridRow(serverRecord: any, entityName: string, gridColumns: string[]) {
        const row: any = {};

        for (const column of gridColumns) {
            if (!this.attributeList.includes(column)) return;

            if (serverRecord[`_${column}_value@Microsoft.Dynamics.CRM.lookuplogicalname`]) {
                if (column === "mserp_fk_vendorapprovedprocurementcategory_id") {
                    const aliasedKey = Object.keys(serverRecord).find(k => k.endsWith("mserp_productcategoryname"));

                    row[column] = {
                        id: serverRecord[`_${column}_value`] ?? null,
                        entityLogicalName: serverRecord[`_${column}_value@Microsoft.Dynamics.CRM.lookuplogicalname`],
                        name: serverRecord[aliasedKey],
                    };
                }
                else {
                    row[column] = {
                        id: serverRecord[`_${column}_value`] ?? null,
                        entityLogicalName: serverRecord[`_${column}_value@Microsoft.Dynamics.CRM.lookuplogicalname`],
                        name: serverRecord[`_${column}_value@OData.Community.Display.V1.FormattedValue`] ?? null
                    };
                }
            } else {
                const rawValue = serverRecord[column];
                const formattedValue = serverRecord[`${column}@OData.Community.Display.V1.FormattedValue`];
                // Dataverse returns decimal/money/integer fields as JS numbers (e.g. 1000).
                // Their FormattedValues are locale-formatted strings (e.g. "1,000.000000") which
                // break parseFloat. If rawValue is a number and formattedValue is a formatted
                // number (only digits, commas, dots), use the raw number directly.
                // This preserves FormattedValue for picklist labels (e.g. "Active"), booleans, etc.
                const fvIsFormattedNumber = typeof formattedValue === "string" && /^-?[\d,.]+$/.test(formattedValue);
                row[column] = (typeof rawValue === "number" && !isNaN(rawValue) && fvIsFormattedNumber)
                    ? rawValue
                    : formattedValue ?? rawValue ?? null;
            }
        }

        row[entityName + "id"] = serverRecord[entityName + "id"];

        return row;
    }

    /**
     * Extracts the entity name from a provided FetchXML query.
     * @param {string} fetchXml - The FetchXML query string.
     * @returns {string} The extracted entity name.
     */
    private getEntityNameFromFetchXml(fetchXml: string) {
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(fetchXml, "text/xml");

        const entity = xmlDoc.getElementsByTagName("entity");

        const entityName = entity[0].getAttribute("name");

        return entityName!;
    }

    /**
     * Adds filtering based on the PO Header.
     * @param {string} headerLogicalName - The logical name of the header
     * @param {string} headerId - The Guid of the Header.
     * @param {string} fetchXml - The FetchXML query string.
     * @returns {string} The finalized fetchxml.
     */
    private addPOHeaderFilterFetchXml(
        headerLogicalName: string,
        headerId: string,
        fetchXml: string,
        cacheField: string
    ) {
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(fetchXml, "text/xml");
        const entityNode = xmlDoc.getElementsByTagName("entity")[0];

        if (entityNode) {
            let filterNode: Element = entityNode.getElementsByTagName("filter")[0];

            if (!filterNode) {
                filterNode = xmlDoc.createElement("filter");
                filterNode.setAttribute("type", "and");
                entityNode.appendChild(filterNode);
            }

            const conditionNode = xmlDoc.createElement("condition");
            conditionNode.setAttribute("attribute", headerLogicalName.toLowerCase());
            conditionNode.setAttribute("operator", "eq");
            conditionNode.setAttribute("value", headerId);

            const filterOr = xmlDoc.createElement("filter");
            if (cacheField && cacheField !== "") {
                filterOr.setAttribute("type", "or");

                const conditionNode1 = xmlDoc.createElement("condition");
                conditionNode1.setAttribute("attribute", cacheField);
                conditionNode1.setAttribute("operator", "le");
                conditionNode1.setAttribute("value", new Date().toISOString());
                const conditionNode2 = xmlDoc.createElement("condition");
                conditionNode2.setAttribute("attribute", cacheField);
                conditionNode2.setAttribute("operator", "ge");
                conditionNode2.setAttribute("value", new Date().toISOString());
                filterOr.appendChild(conditionNode1);
                filterOr.appendChild(conditionNode2);
            }

            filterNode.appendChild(conditionNode);

            if (cacheField && cacheField !== "") filterNode.appendChild(filterOr);

            const attributeNodes = xmlDoc.getElementsByTagName("attribute");
            Array.from(attributeNodes).forEach(attr => this.attributeList.push(attr.getAttribute("name")));
            Array.from(attributeNodes)?.forEach(attr => attr.remove());
        }

        const serializer = new XMLSerializer();
        return serializer.serializeToString(xmlDoc);
    }

    /**
     * Adds filtering based on the searchable columns.
     *
     * Searchable columns are defined in the controls properties
     * @param {string} searchText - The search text.
     * @param {string[]} searchableColumns - The array of searchable columns.
     * @returns {string} The finalized fetchxml.
     */
    private addFiltersOnSearch(
        searchText: string,
        attributesMetadata: any[],
        searchableColumns: string[],
        fetchXml: string
    ) {
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(fetchXml, "text/xml");
        const entityNode = xmlDoc.getElementsByTagName("entity")[0];

        if (!entityNode) return fetchXml;

        const filterNode = xmlDoc.createElement("filter");
        filterNode.setAttribute("type", "or");

        for (const attr of attributesMetadata) {
            const logicalName = attr.LogicalName;
            if (!searchableColumns.includes(logicalName)) continue;

            const dataType = attr.AttributeType.toLowerCase();

            const isVirtualField = attr.AttributeType?.toLowerCase() === "virtual" && attr.SourceType > 0;
            const cachedVirtualFields = [
                "mserp_concatenatedproductdimensions",
                "mserp_invoicestatus",
                "mserp_availabletoinvoice",
                "mserp_outofdate"
            ];

            console.log("Logical name: " + logicalName);
            if (isVirtualField || cachedVirtualFields.includes(logicalName?.toLowerCase())) {
                continue;
            }

            const condition = xmlDoc.createElement("condition");
            condition.setAttribute("attribute", logicalName);

            switch (dataType) {
                case "integer":
                case "int":
                    const intVal = parseInt(searchText.trim());
                    if (!isNaN(intVal)) {
                        condition.setAttribute("operator", "eq");
                        condition.setAttribute("value", intVal.toString());
                        filterNode.appendChild(condition);
                    }
                    break;

                case "decimal":
                case "double":
                case "money":
                    const floatVal = parseFloat(searchText.trim());
                    if (!isNaN(floatVal)) {
                        condition.setAttribute("operator", "eq");
                        condition.setAttribute("value", floatVal.toString());
                        filterNode.appendChild(condition);
                    }
                    break;

                case "datetime":
                case "date":
                    const dateVal = new Date(searchText.trim());
                    if (!isNaN(dateVal.getTime()) && dateVal.getFullYear() >= 1800) {
                        condition.setAttribute("operator", "on");
                        condition.setAttribute("value", dateVal.toISOString().split("T")[0]);
                        filterNode.appendChild(condition);
                    }
                    break;

                case "boolean":
                    if (["true", "false", "yes", "no"].includes(searchText.toLowerCase())) {
                        const boolVal = ["true", "yes"].includes(searchText.toLowerCase()) ? "1" : "0";
                        condition.setAttribute("operator", "eq");
                        condition.setAttribute("value", boolVal);
                        filterNode.appendChild(condition);
                    }
                    break;

                case "picklist":
                case "status":
                case "state":
                case "optionset":
                    const option = LineStatusOptions.find(opt => opt.label.toLowerCase() === searchText.toLowerCase());
                    if (!option) continue;
                    condition.setAttribute("operator", "eq");
                    condition.setAttribute("value", option.value.toString());
                    filterNode.appendChild(condition);
                    break;

                case "lookup":
                    // Skip
                    break;

                default:
                    condition.setAttribute("operator", "like");
                    condition.setAttribute("value", `%${searchText}%`);
                    filterNode.appendChild(condition);
                    break;
            }
        }

        entityNode.appendChild(filterNode);
        const attributeNodes = Array.from(entityNode.getElementsByTagName("attribute"));
        attributeNodes?.forEach(attr => attr.remove());

        return new XMLSerializer().serializeToString(xmlDoc);
    }
}