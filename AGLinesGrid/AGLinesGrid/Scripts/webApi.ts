import { IInputs } from "../generated/ManifestTypes";
import {
    AttributeDTO,
    EntityDTO,
    LineSplitDeliveryAction,
    LineStatus,
    LocalizedDTO,
    NoYes,
    POLineActionButtons,
    RecordsWebApi
} from "./types";

/**
 * @file webApi.ts
 * @description Provides utility functions for interacting with a web API.
 */

/**
 * Converts a singular noun to its plural form.
 *
 * @param {string} str - The singular noun.
 * @returns {string} The pluralized form of the noun.
 */
export function _getPlural(str: string) {
    if (str.endsWith("s") || str.endsWith("x")) str += "es";
    else if (str.endsWith("y")) str = str.substr(0, str.length - 1) + "ies";
    else str += "s";

    return str;
}

/**
 * Performs an AJAX request with token authentication.
 *
 * @param {any} ajaxOptions - The AJAX request options.
 */
export function _safeAjax(ajaxOptions: any) {
    //@ts-ignore
    const deferredAjax = $.Deferred();

    //@ts-ignore
    shell
        .getTokenDeferred()
        //@ts-ignore
        .done(function (token) {
            // add headers for AJAX
            if (!ajaxOptions.headers) {
                //@ts-ignore
                $.extend(ajaxOptions, {
                    headers: {
                        __RequestVerificationToken: token
                    }
                });
            } else {
                ajaxOptions.headers["__RequestVerificationToken"] = token;
            }
            //@ts-ignore
            $.ajax(ajaxOptions)
                .done(function (data: any, textStatus: string, jqXHR: string) {
                    //@ts-ignore
                    validateLoginSession(data, textStatus, jqXHR, deferredAjax.resolve);
                })
                .fail(deferredAjax.reject); //AJAX
        })
        .fail(function () {
            //@ts-ignore
            deferredAjax.rejectWith(this, arguments); // on token failure pass the token AJAX and args
        });

    return deferredAjax.promise();
}

/**
 * Retrieves multiple records from the API for a given entity.
 *
 * @param {string} entityName - The name of the entity to retrieve data from.
 * @param {string} query - The query parameters for filtering data.
 * @param {ComponentFramework.WebApi} contextWebApi - Controls context webapi
 * @returns {Promise<any>} A promise resolving with the retrieved data.
 */
export async function sendRequest(
    entityName: string,
    query: string,
    singleRetrieve: boolean,
    contextWebApi: ComponentFramework.WebApi
): Promise<any> {
    return new Promise((resolve, reject) => {
        if (document.getElementById("antiforgerytoken") !== null) {
            _safeAjax({
                type: "GET",
                url: `/_api/${_getPlural(entityName)}${singleRetrieve ? `(${query})` : `${query}`}`,
                contentType: "application/json",
                success: function (results: any, status: any, xhr: any) {
                    resolve(results);
                },
                error: function (xhr: any, ajaxOptions: any, thrownError: any) {
                    reject(
                        xhr.responseJSON
                            ? xhr.responseJSON?.error?.innererror
                                ? xhr.responseJSON?.error?.innererror.message
                                : xhr.responseJSON.error.message
                            : thrownError
                    );
                }
            });
        } else {
            if (singleRetrieve)
                contextWebApi.retrieveRecord(entityName, query).then(
                    function success(results) {
                        resolve(results);
                    },
                    function (error) {
                        reject(error);
                    }
                );
            else
                contextWebApi.retrieveMultipleRecords(entityName, query).then(
                    function success(results) {
                        resolve(results);
                    },
                    function (error) {
                        reject(error);
                    }
                );
        }
    });
}

/**
 * Retrieves a single record from the API for a given entity.
 *
 * @param {string} entityName - The name of the entity.
 * @param {string} id - The unique identifier of the record.
 * @param {ComponentFramework.WebApi} contextWebApi - Controls context webapi
 * @returns {Promise<any>} A promise resolving with the retrieved record.
 */
export async function retrieve(entityName: string, id: string, contextWebApi: ComponentFramework.WebApi) {
    return await sendRequest(entityName, id, true, contextWebApi);
}

/**
 * Retrieves multiple records from the API for a given entity.
 *
 * @param {string} entityName - The name of the entity to retrieve data from.
 * @param {string} query - The query parameters for filtering data.
 * @param {ComponentFramework.WebApi} contextWebApi - Controls context webapi
 * @returns {Promise<any>} A promise resolving with the retrieved data.
 */
export async function retrieveMultiple(
    entityName: string,
    query: string,
    contextWebApi: ComponentFramework.WebApi
): Promise<any> {
    return await sendRequest(entityName, query, false, contextWebApi);
}

/**
 * Updates a record in the API for a given entity.
 *
 * @param {string} entityName - The name of the entity.
 * @param {string} id - The unique identifier of the record.
 * @param {any} data - The data to update the record with.
 * @returns {Promise<void>} A promise resolving when the update is complete.
 */
export async function updateRecord(
    entityName: string,
    id: string,
    data: any,
    contextWebApi: ComponentFramework.WebApi
): Promise<any> {
    return new Promise((resolve, reject) => {
        if (document.getElementById("antiforgerytoken") !== null) {
            _safeAjax({
                type: "PATCH",
                url: `/_api/${_getPlural(entityName)}(${id})`,
                data: JSON.stringify(data),
                contentType: "application/json",
                success: function (results: any, status: any, xhr: any) {
                    resolve(results);
                },
                error: function (xhr: any, ajaxOptions: any, thrownError: any) {
                    reject(
                        xhr.responseJSON
                            ? xhr.responseJSON?.error?.innererror
                                ? xhr.responseJSON?.error?.innererror.message
                                : xhr.responseJSON.error.message
                            : thrownError
                    );
                }
            });
        } else {
            contextWebApi.updateRecord(entityName, id, data).then(
                function success(results) {
                    resolve(results);
                },
                function (error) {
                    reject(error);
                }
            );
        }
    });
}

/**
 * Updates a record in the API for a given entity.
 *
 * @param {string} entityName - The name of the entity.
 * @param {string} id - The unique identifier of the record.
 * @param {any} data - The data to update the record with.
 * @returns {Promise<void>} A promise resolving when the update is complete.
 */
export async function createRecord(
    entityName: string,
    data: any,
    contextWebApi: ComponentFramework.WebApi,
    headerEntityName: string,
    headerEntityId: string
): Promise<any> {
    if (!data.hasOwnProperty("mserp_dataareaid")) {
        const companyCode = await getCompanyCode(headerEntityName, headerEntityId, contextWebApi);
        data.mserp_dataareaid = companyCode;
    }
    return new Promise((resolve, reject) => {
        if (document.getElementById("antiforgerytoken") !== null) {
            _safeAjax({
                type: "POST",
                url: `/_api/${_getPlural(entityName)}`,
                data: JSON.stringify(data),
                contentType: "application/json",
                success: function (results: any, status: any, xhr: any) {
                    resolve(results);
                },
                error: function (xhr: any, ajaxOptions: any, thrownError: any) {
                    reject(
                        xhr.responseJSON
                            ? xhr.responseJSON?.error?.innererror
                                ? xhr.responseJSON?.error?.innererror.message
                                : xhr.responseJSON.error?.message
                            : thrownError
                    );
                }
            });
        } else {
            contextWebApi.createRecord(entityName, data).then(
                function success(results) {
                    resolve(results);
                },
                function (error) {
                    reject(error);
                }
            );
        }
    });
}

async function getCompanyCode(
    primaryEntityName: string,
    primaryEntityId: string,
    contextWebApi: ComponentFramework.WebApi
) {
    const companyFieldName = "mserp_dataareaid";
    const companyEl = document.getElementById(companyFieldName) as HTMLInputElement;
    if (companyEl) {
        return companyEl.value;
    } else {
        const requestOp = async (): Promise<string | null> => {
            return new Promise((resolve, reject) => {
                if (document.getElementById("antiforgerytoken") !== null) {
                    _safeAjax({
                        type: "GET",
                        url: `/_api/${_getPlural(primaryEntityName)}(${primaryEntityId})?$select=mserp_dataareaid`,
                        contentType: "application/json",
                        success: function (data: any, status: any, xhr: any) {
                            if (!data) resolve(null);

                            resolve(
                                data.mserp_dataareaid ?? null
                            );
                        },
                        error: function (xhr: any, ajaxOptions: any, thrownError: any) {
                            reject(
                                xhr.responseJSON
                                    ? xhr.responseJSON?.error?.innererror
                                        ? xhr.responseJSON?.error?.innererror.message
                                        : xhr.responseJSON.error.message
                                    : thrownError
                            );
                        }
                    });
                } else {
                    contextWebApi.retrieveRecord(primaryEntityName, primaryEntityId, "?$select=mserp_dataareaid").then(
                        function success(result) {
                            if (!result) resolve(null);
                            resolve(
                                result.mserp_dataareaid ?? null
                            );
                        },
                        function (error) {
                            reject(error);
                        }
                    );
                }
            });
        };

        return await requestOp();
    }
}

/**
 * Updates a record in the API for a given entity.
 *
 * @param {string} entityName - The name of the entity.
 * @param {string} id - The unique identifier of the record.
 * @param {any} data - The data to update the record with.
 * @returns {Promise<void>} A promise resolving when the update is complete.
 */
export async function executeAction(
    gridRowActionName: string,
    entityName: string,
    id: string,
    actionName: string,
    contextWebApi: ComponentFramework.WebApi,
    data?: LineSplitDeliveryAction | any
): Promise<any> {
    const updateData = {
        mserp_actionname: actionName,
        ...(data ? { mserp_actionparameters: data } : {})
    };

    if (document.getElementById("antiforgerytoken") !== null) {
        if (gridRowActionName === POLineActionButtons.SplitDelivery)
            return new Promise((resolve, reject) => {
                _safeAjax({
                    type: "PATCH",
                    url: `/_api/${_getPlural(entityName)}(${id})`,
                    data: JSON.stringify(updateData),
                    contentType: "application/json",
                    success: function (data: any, status: any, xhr: any) {
                        resolve(data);
                    },
                    error: function (xhr: any, ajaxOptions: any, thrownError: any) {
                        reject(
                            xhr.responseJSON
                                ? xhr.responseJSON?.error?.innererror
                                    ? xhr.responseJSON?.error?.innererror.message
                                    : xhr.responseJSON.error.message
                                : thrownError
                        );
                    }
                });
            });

        return new Promise((resolve, reject) => {
            _safeAjax({
                type: "PUT",
                url: `/_api/${_getPlural(entityName)}(${id})/mserp_actionname`,
                data: JSON.stringify({ value: actionName }),
                contentType: "application/json",
                success: function (data: any, status: any, xhr: any) {
                    resolve(data);
                },
                error: function (xhr: any, ajaxOptions: any, thrownError: any) {
                    reject(
                        xhr.responseJSON
                            ? xhr.responseJSON?.error?.innererror
                                ? xhr.responseJSON?.error?.innererror.message
                                : xhr.responseJSON.error.message
                            : thrownError
                    );
                }
            });
        });
    } else {
        return new Promise((resolve, reject) => {
            contextWebApi.updateRecord(entityName, id, updateData).then(
                function success(results) {
                    resolve(results);
                },
                function (error) {
                    reject(error);
                }
            );
        });
    }
}

/**
 * Deletes a record in the API for a given entity.
 *
 * @param {string} entityName - The name of the entity.
 * @param {string} id - The unique identifier of the record.
 * @param {any} data - The data to update the record with.
 * @returns {Promise<void>} A promise resolving when the update is complete.
 */
export async function deleteRecord(
    entityName: string,
    id: string,
    contextWebApi: ComponentFramework.WebApi
): Promise<any> {
    return new Promise((resolve, reject) => {
        if (document.getElementById("antiforgerytoken") !== null) {
            _safeAjax({
                type: "DELETE",
                url: `/_api/${_getPlural(entityName)}(${id})`,
                contentType: "application/json",
                success: function (result: any, status: any, xhr: any) {
                    resolve(result);
                },
                error: function (xhr: any, ajaxOptions: any, thrownError: any) {
                    reject(
                        xhr.responseJSON
                            ? xhr.responseJSON?.error?.innererror
                                ? xhr.responseJSON?.error?.innererror.message
                                : xhr.responseJSON.error.message
                            : thrownError
                    );
                }
            });
        } else {
            contextWebApi.deleteRecord(entityName, id).then(
                function success(results) {
                    resolve(results);
                },
                function (error) {
                    reject(error);
                }
            );
        }
    });
}

/**
 * Returns a list of suggested lookup values for a given filter text.
 * @param entityType Entity type name.
 * @param filterText A string by which the displayField will be filtered
 * @returns Suggested lookup values.
 */
export async function getLookupValues(
    fetchXml: string,
    entityType: string,
    displayField: string,
    secondaryDisplayField: string | null,
    currentPage: number,
    _webApi: ComponentFramework.WebApi
): Promise<RecordsWebApi> {
    const recordsWebApi: RecordsWebApi = {
        records: [],
        hasMoreRecords: false,
        pagingCookie: undefined,
        pageNumber: currentPage
    };

    try {
        const data = await retrieveMultiple(
            entityType,
            "?fetchXml=" + encodeURIComponent(compressFetchXml(fetchXml)),
            _webApi
        );
        const entities = data.value || data.entities || [];
        const values = entities.map((r: ComponentFramework.WebApi.Entity) => {
            const val: any = {
                entityType: entityType,
                id: r[`${entityType}id`],
                name: r[`${displayField}`] || ""
            };

            val[displayField] = r[`${displayField}`] || "";

            if (secondaryDisplayField) {
                val[secondaryDisplayField] = r[`${secondaryDisplayField}`] || "";
            }

            return val;
        });

        recordsWebApi.records = values;

        if (data["@Microsoft.Dynamics.CRM.fetchxmlpagingcookie"] && data["@Microsoft.Dynamics.CRM.morerecords"]) {
            recordsWebApi.hasMoreRecords = data["@Microsoft.Dynamics.CRM.morerecords"];
            recordsWebApi.pagingCookie = data["@Microsoft.Dynamics.CRM.fetchxmlpagingcookie"] ?? undefined;
            recordsWebApi.pageNumber = extractPageNumber(data["@Microsoft.Dynamics.CRM.fetchxmlpagingcookie"]);
        }

        return recordsWebApi;
    } catch (e) {
        alert(e);
    }
}

/**
 *
 * @param entityType The entity type
 * @param displayField Display column that will show
 * @param secondaryDisplayField Secondary column shown in the dropdown table
 * @param filterFieldId Filtering field if any
 * @param filterFieldRelationColumn Filtering relation column
 * @param filterText Filtering text
 * @param saveTextValue Should the value be saved in the text field
 * @param saveTextFieldName Which field should the go into the text field
 * @param page Page
 * @param isPaging Is it paging
 * @param vendorId The Vendor Guid (used for Invoices)
 * @param pagingCookie Paging cookie
 * @returns Compressed FetchXml
 */
export function getFetchXml(
    entityType: string,
    displayField: string,
    secondaryDisplayField: string | null,
    filterFieldId: string | null,
    filterFieldRelationColumn: string | null,
    filterText: string,
    page: number,
    isPaging: boolean = false,
    vendorId: string | null,
    pagingCookie?: string
) {
    const cookie = isPaging ? `paging-cookie="` + extractAndProcessPagingCookie(pagingCookie) + `"` : "";

    return compressFetchXml(`<fetch version="1.0" output-format="xml-platform" mapping="logical" count="10" page="${page}" ${cookie}>
                  <entity name="${entityType}">
                    <attribute name="${entityType}id" />
                    <attribute name="${displayField}" />
                    ${secondaryDisplayField && `<attribute name="${secondaryDisplayField}" />`}
                    <order attribute="${displayField}" descending="false" />
                    ${
                        filterText
                            ? `<filter type='or'>
                      <condition attribute="${displayField}" operator="like" value="%${filterText}%" />
                      ${
                          secondaryDisplayField &&
                          `<condition attribute="${secondaryDisplayField}" operator="like" value="%${filterText}%" />`
                      }
                    </filter>`
                            : ""
                    }
                    ${
                        filterFieldId && filterFieldId !== "" && filterFieldRelationColumn
                            ? `<filter>
                            <condition attribute="${filterFieldRelationColumn}" operator="eq" value="${filterFieldId}" />
                        </filter>`
                            : ""
                    }
                    ${
                        entityType === "mserp_vrmprocurementproductcategoryvendorassignmententity" && vendorId !== null
                            ? `<filter type="and">
                            <condition attribute="mserp_validto" operator="ge" value="${new Date().toISOString()}" />
                        </filter>
                        <filter type="and">
                            <condition attribute="mserp_fk_vendort_id" operator="eq" value="${vendorId}" />
                        </filter>`
                            : ""
                    }
                  </entity>
                </fetch>`);
}

export async function getSplitDeliveryData(headerId: string, lineNumber: number, _webApi: ComponentFramework.WebApi) {
    if (!headerId || !lineNumber) return [];

    const fetchXml = `<fetch version="1.0" output-format="xml-platform" mapping="logical" no-lock="true">
                        <entity name="mserp_vrmpurchaseorderresponselineentity">
                            <attribute name="mserp_confirmeddlv" />
                            <attribute name="mserp_deliverydate" />
                            <attribute name="mserp_purchqty" />
                            <attribute name="mserp_linestate" />
                            <filter>
                                <condition attribute="mserp_fk_responseheader_id" operator="eq" value="${headerId}" />
                                <condition attribute="mserp_linenumber" operator="eq" value="${lineNumber}" />
                            </filter>
                        </entity>
                       </fetch>`;

    try {
        const data = await retrieveMultiple(
            "mserp_vrmpurchaseorderresponselineentity",
            "?fetchXml=" + encodeURIComponent(compressFetchXml(fetchXml)),
            _webApi
        );

        const entities = data.value || data.entities || [];
        const values = entities
            .filter((r: ComponentFramework.WebApi.Entity) => r?.mserp_linestate === LineStatus.ScheduleLine)
            .map((r: ComponentFramework.WebApi.Entity, index: number) => {
                const val: any = {
                    id: index + 1,
                    mserp_purchqty: r?.mserp_purchqty,
                    mserp_deliverydate:
                        r?.mserp_deliverydate && !isNaN(new Date(r?.mserp_deliverydate).getTime())
                            ? new Date(r?.mserp_deliverydate)
                            : null,
                    mserp_confirmeddlv:
                        r?.mserp_confirmeddlv && !isNaN(new Date(r?.mserp_confirmeddlv).getTime())
                            ? new Date(r?.mserp_confirmeddlv)
                            : null
                };

                return val;
            });

        return values;
    } catch (e) {
        alert(e);
        return [];
    }
}

/**
 * Compresses the query removing whitespaces,newlines and trims it
 */
function compressFetchXml(fetchXml: string): string {
    return fetchXml
        .replace(/(\r\n|\n|\r)/gm, "")
        .replace(/\s{2,}/g, " ")
        .trim();
}

/**
 * Extracts the paging cookie from the returned results
 */
function extractAndProcessPagingCookie(pagingCookie: string | undefined): string {
    if (!pagingCookie) return "";

    // Extract only the value inside `pagingcookie="..."` using regex
    const match = pagingCookie.match(/pagingcookie="([^"]+)"/);

    if (!match || match.length < 2) {
        console.warn("Paging cookie not found or invalid.");
        return "";
    }

    // Extracted raw paging cookie value
    const rawCookie = match[1];

    // Double decode the extracted value
    const decodedCookie = decodeURIComponent(decodeURIComponent(rawCookie));

    // Replace special characters with XML-safe entities
    return decodedCookie.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

/**
 * Extracts the next page number from the paging cookie
 */
function extractPageNumber(pagingCookie: string): number | null {
    if (!pagingCookie) return null;

    const match = pagingCookie.match(/pagenumber="(\d+)"/);

    if (match && match[1]) {
        return parseInt(match[1], 10);
    }

    console.warn("Page number not found in paging cookie.");
    return null;
}

export function normalizeCheckboxValue(val: any): NoYes {
    if (val === true || val === "Yes" || val === 1 || val === NoYes.Yes) return NoYes.Yes;
    return NoYes.No;
}

/**
 * Checks whether the Invoice has a Purchase Order.
 * @param primaryEntityName The primary entity name to take the id from
 * @param primaryEntityId The primary entity id
 * @param columnName Column name of the id to be retrieved
 * @param contextWebApi Context web api
 * @returns
 */
export async function checkInvoicesPurchaseOrder(
    primaryEntityName: string,
    primaryEntityId: string,
    columnName: string = "mserp_fk_purchaseorderheader_id",
    contextWebApi: ComponentFramework.WebApi
) {
    const purchaseOrder = document.getElementById(columnName) as HTMLInputElement;
    if (purchaseOrder) {
        return purchaseOrder.value !== null && purchaseOrder.value !== "" ? false : true;
    } else {
        const requestOp = async (): Promise<boolean | null> => {
            return new Promise((resolve, reject) => {
                if (document.getElementById("antiforgerytoken") !== null) {
                    _safeAjax({
                        type: "GET",
                        url: `/_api/${_getPlural(primaryEntityName)}(${primaryEntityId})?$select=_${columnName}_value`,
                        contentType: "application/json",
                        success: function (data: any, status: any, xhr: any) {
                            if (!data) resolve(true);

                            resolve(data[`_${columnName}_value`] !== null ? false : true);
                        },
                        error: function (xhr: any, ajaxOptions: any, thrownError: any) {
                            reject(
                                xhr.responseJSON
                                    ? xhr.responseJSON?.error?.innererror
                                        ? xhr.responseJSON?.error?.innererror.message
                                        : xhr.responseJSON.error.message
                                    : thrownError
                            );
                        }
                    });
                } else {
                    contextWebApi
                        .retrieveRecord(primaryEntityName, primaryEntityId, `?$select=_${columnName}_value`)
                        .then(
                            function success(result) {
                                if (!result) resolve(true);
                                resolve(result[`_${columnName}_value`] !== null ? false : true);
                            },
                            function (error) {
                                reject(error);
                            }
                        );
                }
            });
        };

        return await requestOp();
    }
}

/**
 * Checks whether the Invoice has a Purchase Order.
 * @param primaryEntityName The primary entity name to take the id from
 * @param primaryEntityId The primary entity id
 * @param columnName Column name of the id to be retrieved
 * @param contextWebApi Context web api
 * @returns
 */
export async function getVendorForInvoice(
    primaryEntityName: string,
    primaryEntityId: string,
    columnName: string = "mserp_fk_vendor_id",
    contextWebApi: ComponentFramework.WebApi
) {
    const vendorId = document.getElementById(columnName) as HTMLInputElement;
    if (vendorId) {
        return vendorId.value;
    } else {
        const requestOp = async (): Promise<boolean | null> => {
            return new Promise((resolve, reject) => {
                if (document.getElementById("antiforgerytoken") !== null) {
                    _safeAjax({
                        type: "GET",
                        url: `/_api/${_getPlural(primaryEntityName)}(${primaryEntityId})?$select=_${columnName}_value`,
                        contentType: "application/json",
                        success: function (data: any, status: any, xhr: any) {
                            if (!data) resolve(null);

                            resolve(data[`_${columnName}_value`] || null);
                        },
                        error: function (xhr: any, ajaxOptions: any, thrownError: any) {
                            reject(
                                xhr.responseJSON
                                    ? xhr.responseJSON?.error?.innererror
                                        ? xhr.responseJSON?.error?.innererror.message
                                        : xhr.responseJSON.error.message
                                    : thrownError
                            );
                        }
                    });
                } else {
                    contextWebApi
                        .retrieveRecord(primaryEntityName, primaryEntityId, `?$select=_${columnName}_value`)
                        .then(
                            function success(result) {
                                if (!result) resolve(null);
                                resolve(result[`_${columnName}_value`] || null);
                            },
                            function (error) {
                                reject(error);
                            }
                        );
                }
            });
        };

        return await requestOp();
    }
}

// Cached entity definitions on Portal
const entityDefsCache: Map<string, EntityDTO> = new Map();

function getEntityDefinitionsBundle(): { entities: EntityDTO[] } | null {
    return (
        (window as any).msdyn?.Portal?.EntityDefinitions ??
        (parent.window as any).msdyn?.Portal?.EntityDefinitions ??
        null
    );
}

function hydrateCacheOnce() {
    if (entityDefsCache.size > 0) return;
    const bundle = getEntityDefinitionsBundle();
    if (!bundle?.entities?.length) return;
    for (const e of bundle.entities) {
        entityDefsCache.set(e.logicalName.toLowerCase(), e);
    }
}

export function getEntityDto(logicalName: string): EntityDTO | null {
    hydrateCacheOnce();
    return entityDefsCache.get(logicalName.toLowerCase()) ?? null;
}

export function getAttributeDto(entity: EntityDTO, attributeLogicalName: string): AttributeDTO | null {
    const attrs = entity.attributes ?? [];
    const found = attrs.find(a => a.logicalName?.toLowerCase() === attributeLogicalName.toLowerCase());
    return found ?? null;
}

/** Returns the lookup targets array or [] if not a lookup/unknown. */
export function getLookupTargetsFromPortalDefs(
    entityLogicalName: string,
    lookupAttributeLogicalName: string
): string[] {
    const ent = getEntityDto(entityLogicalName);
    if (!ent) return [];
    const attr = getAttributeDto(ent, lookupAttributeLogicalName);
    if (!attr) return [];
    if ((attr.type ?? "").toLowerCase() !== "lookup") return [];
    return Array.isArray(attr.targets) ? attr.targets.slice() : [];
}

/** Prefer a specific LCID if provided; otherwise try 1033, else first available. */
function pickLocalizedLabel(
    dict: Record<number, string> | undefined,
    fallback?: string,
    preferredLcid?: number
): string | undefined {
    if (!dict || Object.keys(dict).length === 0) return fallback;
    if (preferredLcid && dict[preferredLcid]) return dict[preferredLcid];
    if (dict[1033]) return dict[1033];
    const first = dict[Number(Object.keys(dict)[0])];
    return first ?? fallback;
}

/** Get entity DTO from the preloaded bundle */
function getLookupEntityDtoFromPortal(lookupEntityType: string): EntityDTO | null {
    return getEntityDto(lookupEntityType);
}

export async function buildLookupEntityColumns(
    context: ComponentFramework.Context<IInputs>,
    lookupEntityType: string | undefined,
    attributes: string[],
    preferredLcid?: number
): Promise<Array<{ key: string; label: string }>> {
    if (!lookupEntityType) {
        throw new Error("Error retrieving Lookup Entity Metadata.");
    }

    // 1) Try portal bundle
    const portalEnt = getLookupEntityDtoFromPortal(lookupEntityType);
    if (portalEnt?.attributes?.length) {
        return attributes.map(col => {
            const attr = portalEnt.attributes!.find(a => a.logicalName?.toLowerCase() === col.toLowerCase());
            const label =
                pickLocalizedLabel(attr?.displayNameLocalized as LocalizedDTO, attr?.displayName, preferredLcid) ?? col;
            return { key: col, label };
        });
    }

    // 2) Fallback to context.utils
    try {
        const lookupEntityMetadata = await context.utils.getEntityMetadata(lookupEntityType, attributes);
        const lookupEntityAttributesMetadata = lookupEntityMetadata?.Attributes.get?.() ?? [];

        return attributes.map(col => {
            const attrMeta = lookupEntityAttributesMetadata.find(
                (m: any) => (m.LogicalName ?? "").toLowerCase() === col.toLowerCase()
            );
            const label = attrMeta?.DisplayName?.UserLocalizedLabel?.Label ?? attrMeta?.DisplayName ?? col;
            return { key: col, label };
        });
    } catch {
        // 3) Last resort: echo keys
        return attributes.map(col => ({ key: col, label: col }));
    }
}

export async function getEntityMetadataUnified(
    context: ComponentFramework.Context<any>,
    entityLogicalName: string,
    attributes?: string[],
    preferredLcid?: number
) {
    const shim = getEntityMetadataFromPortalShim(entityLogicalName, attributes, preferredLcid);
    if (shim) return shim; // portal-backed
    return await context.utils.getEntityMetadata(entityLogicalName, attributes); // live fallback
}

/** Adapter: returns a minimal shim compatible with context.utils.getEntityMetadata(...) */
export function getEntityMetadataFromPortalShim(
    entityLogicalName: string,
    requestedAttributes?: string[],
    preferredLcid?: number
): { Attributes: { get: () => any[] } } | null {
    const ent = getEntityDto(entityLogicalName);
    if (!ent?.attributes?.length) return null;

    const want = (requestedAttributes ?? []).map(a => a.toLowerCase());
    const list = (ent.attributes ?? []).filter(a =>
        want.length ? want.includes((a.logicalName ?? "").toLowerCase()) : true
    );

    const normalized = list.map(a => ({
        LogicalName: a.logicalName,
        SchemaName: a.schemaName,
        DisplayName: {
            UserLocalizedLabel: {
                Label: pickLocalizedLabel(a.displayNameLocalized as any, a.displayName, preferredLcid) ?? a.logicalName
            }
        },
        AttributeType: a.type, // optional, helpful to have
        Targets: Array.isArray(a.targets) ? a.targets.slice() : undefined,
        IsValidForRead: a.isValidForRead // optional
    }));

    return {
        Attributes: {
            get: () => normalized
        }
    };
}
