/**
 * @file types.ts
 * @description Defines TypeScript interfaces and types used throughout the application.
 */

import { CustomCellRendererProps } from "ag-grid-react";

/**
 * Represents a column definition in the grid.
 */
export type ColumnDefinition = {
    /** The unique identifier for the column. */
    field?: string;

    /** The display name of the column. */
    headerName?: string;

    /** Options of the column (undefined if the column is not of the type Picklist, Status, State or Boolean). */
    options?: Option[];

    /** Whether the column can be sorted. */
    sortable?: boolean;

    /** Determines if the column is editable. */
    editable?: boolean;

    /** Whether the column is hidden. */
    hide?: boolean;

    /** The type of data in the column (e.g., "string", "boolean", "datetime"). */
    dataType?: string;

    /** Whether the column can be toggled for visibility. */
    toggleShowColumn?: boolean;

    /** Indicates if toggling visibility is enabled. */
    toggleShowColumnEnabled?: boolean;

    /** The index position of the column. */
    index?: number;

    /** The width of the column in pixels. */
    width?: number;

    /** The original value field name for comparison. Used for the tooltip if the original value is different from the current value. */
    fieldOriginalValue?: string;

    /** The original value field Display Name */
    fieldOriginalDisplayName?: string;

    /** Show tooltip of the original value of the cell. The tooltip will be shown when the value has been changed. */
    showOriginalValueTooltip?: boolean;

    /** Whether the column should be included when updating data in Dataverse */
    allowUpdate?: boolean;

    /** Value to be compared against when doing updates */
    compareColumnWith?: string;

    /** On what Line Statuses should the field be editable */
    editableOnStatus?: string[];

    /** If value should be converted to integer */
    convertToInteger?: boolean;

    /** Is checkbox? */
    showAsCheckbox?: boolean;

    /** Lookup configuration */
    lookupColumns?: string[];

    /** Columns for lookups */
    columns?: { key: string; label: string }[];

    /** If a field should be shown in additional details */
    additionalDetails?: boolean;

    /** Alternate field name */
    alternateField?: string;

    /** Whether the line is editable if alternate */
    editableOnAlternate?: boolean;

    /** Which fields are editable if alternate */
    editableAlternateFields?: string[];
};

/**
 * Properties for configuring the grid component.
 */
export type GridProps = {
    /** Entity name of the line */
    entityName: string;

    /** Entity id of the Purchase Order Response Header */
    headerId: string;

    /** Schema name of the Purchase Order Response Header column (Lookup on the Line) */
    headerSchemaName: string;

    /** Entity name of the Purchase Order Response Header */
    headerEntityName: string;

    /** Whether the grid is in read-only mode. */
    readonly: boolean;

    /** Whether row actions are visible. */
    showRowActions: boolean;

    /** The index up to which columns should be pinned. */
    pinIndex: number;

    /** The custom column definitions. */
    customColumns: ColumnDefinition[];

    /** The actions available for each row. */
    rowActions: GridRowActionProps[];

    /** Function to fetch column definitions. */
    retrieveColumnDefs: (lookups: ColumnDefinition[]) => Promise<ColumnDefinition[]>;

    /** Function to fetch the row data. */
    retrieveRows: (
        isSearching: boolean,
        page: number,
        searchText?: string
    ) => Promise<{ data: any[]; totalRecordCount: number }>;

    /** Function to fetch row data based on the searchbox */
    retrieveFilteredRows: (searchText: string) => Promise<any[]>;

    /** Callback triggered when a row value is changed. */
    onRowValueChanged: (
        id: string,
        data: any,
        isCreate: boolean,
        setLoading: React.Dispatch<React.SetStateAction<boolean>>
    ) => Promise<void>;

    /** Page number */
    pageNumber: number;

    /** The Line Status schema/logical name */
    lineStatusLogicalName: string;

    /** The attributes metadata */
    attributesMetadata: any[];

    /** Localization */
    getString(s: string): string;

    /** Entity Metadata */
    entityMetadata: ComponentFramework.PropertyHelper.EntityMetadata;

    /** Ref of the div holder of the control */
    controlRef: React.MutableRefObject<HTMLDivElement>;

    /** Control Unique Identifier
     * This ID should be the same for Grid and Attachment controls that are related
     * Its used to define what Attachment control should be used when clicking on a line and dispatching event,
     * since there can be multiple controls on a single form
     */
    controlId: string;

    /** WebApi */
    webApi: ComponentFramework.WebApi;

    /** Whether to show No of Comments and Attachments columns */
    showCommentsAndAttachments: boolean;

    /** Comments count field */
    commentsCountField: string;

    /** Attachments count field */
    attachmentsCountField: string;

    /** Alternate id on the line used for bundled attachments */
    alternateLineFieldName: string;

    /** Field to compare when hiding price fields */
    hidePriceField?: string;

    /** Which fields to be hidden based on the specific price field */
    hidePriceFields?: string[];

    /** Whether to show mark all as done button */
    showMarkAllButton?: boolean;

    /** Whether to show the debug panel */
    debugMode?: boolean;

    /** Whether to allow to create new lines */
    allowCreate?: boolean;

    /** Vendors for Supplier Portal */
    vendorId?: string | null;

    /** Header status */
    headerStatus?: number | null;
};

/**
 * Represents an action that can be performed on a grid row.
 */
export type GridRowActionProps = {
    /** The label of the action. */
    displayName: string;

    /** Whether the action is disabled. */
    disabled?: boolean;

    /** The function executed when the action is triggered. */
    onClick: (props: CustomCellRendererProps, data?: any) => Promise<any>;

    /** Logical name of the action */
    logicalName: string;
};

/**
 * Represents an option in a selection list.
 */
export type Option = {
    /** The display label of the option. */
    label?: string;

    /** The value associated with the option. */
    value: any;
};

export type LineWebApi = {
    value: any[];
    "@Microsoft.Dynamics.CRM.totalrecordcount": any;
};

export const StateOptions: Option[] = [
    {
        label: undefined,
        value: undefined
    },
    {
        label: "Active",
        value: 0
    },
    {
        label: "Inactive",
        value: 1
    }
];

export const BooleanOptions: Option[] = [
    {
        label: undefined,
        value: null
    },
    {
        label: "Yes",
        value: true
    },
    {
        label: "No",
        value: false
    }
];

export const LineDoneOptions: Option[] = [
    {
        label: "Yes",
        value: 200000001
    },
    {
        label: "No",
        value: 200000000
    }
];

export enum NoYes {
    Yes = 200000001,
    No = 200000000
}

export enum LineStatus {
    Accepted = 200000000,
    AcceptedWithChanges = 200000001,
    SplitIntoSchedule = 200000002,
    ScheduleLine = 200000003,
    Substituted = 200000004,
    Substitute = 200000005,
    Rejected = 200000006,
    SplitFromSubstitute = 200000007
}

export enum RFQLineType {
    Category = "Category",
    Item = "Item"
}

export const LineStatusOptions: Option[] = [
    { label: "Accepted", value: LineStatus.Accepted },
    { label: "Accepted with changes", value: LineStatus.AcceptedWithChanges },
    { label: "Split into schedule", value: LineStatus.SplitIntoSchedule },
    { label: "Schedule line", value: LineStatus.ScheduleLine },
    { label: "Substituted", value: LineStatus.Substituted },
    { label: "Substitute", value: LineStatus.Substitute },
    { label: "Rejected", value: LineStatus.Rejected },
    { label: "Split from substitute", value: LineStatus.SplitFromSubstitute }
];

export interface FilterConfiguration {
    attributeName: string;
    operator: "eq" | "like" | "on";
    type: string;
    options?: string;
    entityName?: string;
    linkEntity?: {
        name: string;
        from: string;
        to: string;
        linkType: string;
    };
}

/**
 * Value needs to be in string
 * Value: { Quantity: number, ConfirmedReceiptDate: string }[]
 * @example { Quantity: 10.00, ConfirmedReceiptDate: '2025-04-10' }
 */
export interface LineSplitDeliveryAction {
    InputParameters: {
        Name: string;
        Value: string;
    }[];
}

export const LineType = {
    VEPurchaseOrderResponseHeader: "mserp_vrmpurchaseorderresponseheaderentity",
    VEPurchaseOrderResponseLine: "mserp_vrmpurchaseorderresponselineentity",
    VEPurchaseOrderResponseHeaderHistory: "mserp_vrmpurchaseorderresponsearchivedheaderentity",
    VEPurchaseOrderResponseLineHistory: "mserp_vrmpurchaseorderresponsearchivedlineentity",
    VEPurchaseOrderResponseArchivedOriginalOrderLine: "mserp_vrmpurchaseorderresponsearchivedoriginalorderlineentity",
    VEPurchaseOrderResponseOriginalorderLine: "mserp_vrmpurchaseorderresponseoriginalorderlineentity",
    VEPurchaseOrderConfirmationHeader: "mserp_vrmpurchaseorderconfirmationheaderentity",
    VEPurchaseOrderConfirmationLine: "mserp_vrmpurchaseorderconfirmationlineentity",
    VERFQHeader: "mserp_vrmrequestforquotationreplyheaderentity",
    VERFQLine: "mserp_vrmrequestforquotationlineentity",
    VERFQReplyHeader: "mserp_vrmrequestforquotationreplyheaderentity",
    VERFQReplyLine: "mserp_vrmrequestforquotationreplylineentity",
    VERequestNewCategory: "mserp_vrmrequestnewcategoryheaderentity",
    VEPendingVendorInvoiceHeader: "mserp_vrmpendingvendorinvoiceheaderdetailsentity",
    VEPendingVendorInvoiceLine: "mserp_vrmpendingvendorinvoicelinedetailsentity",
    VEPOConsignmentInventoryHeader: "mserp_vrmpurchaseorderconsignmentinventoryheaderentity",
    VEPOConsignmentInventoryLine: "mserp_vrmpurchaseorderconsignmentinventorylineentity",
    VERFQAmendment: "mserp_vrmrfqamendmententity",
    VendorBankAccount: "mserp_vrmvendorbankaccountentity",
    VendorCertification: "mserp_vrmvendorcertificationentity"
};

export const POLineActionButtons = {
    RejectLine: "REJECT_LINE",
    SplitDelivery: "SPLIT_DELIVERY",
    Substitute: "SUBSTITUTE_LINE",
    DiscardChanges: "DISCARD_CHANGES"
};

export const RFQLineActionButtons = {
    AddAlternativeLine: "ADD_ALTERNATIVE_LINE",
    RemoveAlternativeLine: "REMOVE_ALTERNATIVE_LINE",
    ResetLine: "RESET_LINE"
};

export const PendingVendorLineActionButtons = {
    DeleteLine: "DELETE_LINE"
};

export const VirtualEntityActions = {
    [LineType.VEPurchaseOrderResponseLine]: {
        Reject: "reject",
        Discard: "revert",
        Substitute: "substitute",
        SplitDelivery: "split",
        MarkAsDone: "MarkLinesDone"
    },
    [LineType.VERFQReplyLine]: {
        AddAlternativeLine: "AddAlternativeLine",
        RemoveAlternativeLine: "RemoveAlternativeLine",
        ResetLine: "reset",
        MarkAsDone: "MarkLinesDone"
    }
};

export const AttachmentEntityName = {
    [LineType.VEPurchaseOrderResponseHeader]: "mserp_vrmpurchaseorderresponseheaderattachmententity",
    [LineType.VEPurchaseOrderResponseLine]: "mserp_vrmpurchaseorderresponselineattachmententity",
    [LineType.VEPurchaseOrderResponseHeaderHistory]: "mserp_vrmpurchaseorderresponseheaderhistoryattachmententity",
    [LineType.VEPurchaseOrderResponseLineHistory]: "mserp_vrmpurchaseorderresponselinehistoryattachmententity",
    [LineType.VEPurchaseOrderResponseArchivedOriginalOrderLine]: "mserp_vrmpurchaseorderresponselinehistoryattachmententity",
    [LineType.VEPurchaseOrderConfirmationLine]: "mserp_vrmpurchaseorderlineattachmententity",
    [LineType.VEPurchaseOrderResponseOriginalorderLine]: "mserp_vrmpurchaseorderresponselineattachmententity",
    [LineType.VERFQHeader]: "mserp_vrmrfqheaderattachmententity",
    [LineType.VERFQLine]: "mserp_vrmrfqlineattachmententity",
    [LineType.VERFQReplyHeader]: "mserp_vrmrfqreplyheaderattachmententity",
    [LineType.VERFQReplyLine]: "mserp_vrmrfqreplylineattachmententity",
    [LineType.VERequestNewCategory]: "mserp_vrmrequestnewcategoryattachmententity",
    [LineType.VEPendingVendorInvoiceHeader]: "mserp_vrmpendingvendorinvoiceheaderattachmentdetailsentity",
    [LineType.VEPendingVendorInvoiceLine]: "mserp_vrmpendingvendorinvoicelineattachmentdetailsentity",
    [LineType.VEPOConsignmentInventoryHeader]: "mserp_vrmpurchaseorderheaderattachmententity",
    [LineType.VEPOConsignmentInventoryLine]: "mserp_vrmpurchaseorderlineattachmententity",
    [LineType.VERFQAmendment]: "mserp_vrmrfqamendmentattachmententity",
    [LineType.VendorBankAccount]: "mserp_vrmvendorbankaccountattachmententity",
    [LineType.VendorCertification]: "mserp_vrmvendorcertificationattachmententity"
};

export const AttachmentFieldName = {
    Lookup: {
        [LineType.VEPurchaseOrderResponseLine]: "mserp_FK_VRMPurchaseOrderResponseLine_id",
        [LineType.VEPurchaseOrderResponseHeaderHistory]: "mserp_FK_ResponseHeader_id",
        [LineType.VEPurchaseOrderResponseLineHistory]: "mserp_FK_ResponseLine_id",
        [LineType.VEPurchaseOrderResponseArchivedOriginalOrderLine]: "mserp_FK_ResponseLine_id",
        [LineType.VEPurchaseOrderConfirmationLine]: "mserp_FK_ConfirmationOrderLine_id",
        [LineType.VEPurchaseOrderResponseOriginalorderLine]: "mserp_FK_ResponseOriginalOrderLine_id",
        [LineType.VERFQHeader]: "mserp_FK_RFQHeader_id",
        [LineType.VERFQLine]: "mserp_FK_RFQLine_id",
        [LineType.VERFQReplyHeader]: "mserp_FK_VRMRequestForQuotationReplyHeader_id",
        [LineType.VERFQReplyLine]: "mserp_FK_VRMRequestForQuotationReplyLine_id",
        [LineType.VERequestNewCategory]: "mserp_FK_RequestNewCategoryHeader_id",
        [LineType.VEPendingVendorInvoiceHeader]: "mserp_FK_PendingVendorInvoiceHeader_id",
        [LineType.VEPendingVendorInvoiceLine]: "mserp_FK_PendingVendorInvoiceLine_id",
        [LineType.VEPOConsignmentInventoryHeader]: "mserp_FK_PurchaseOrderHeader_id",
        [LineType.VEPOConsignmentInventoryLine]: "mserp_FK_PurchaseOrderLine_id",
        [LineType.VERFQAmendment]: "mserp_FK_RFQAmendment_id",
        [LineType.VendorBankAccount]: "mserp_FK_VendorBankAccount_id",
        [LineType.VendorCertification]: "mserp_FK_VendorCertification_id"
    },
    LineNumber: {
        [LineType.VEPurchaseOrderResponseLine]: "mserp_linenumber",
        [LineType.VEPurchaseOrderResponseOriginalorderLine]: "mserp_linenumber",
        [LineType.VERFQLine]: "mserp_rfqreplylinenum",
        [LineType.VERFQReplyLine]: "mserp_rfqreplylinenum",
        [LineType.VEPendingVendorInvoiceLine]: "mserp_linenumber",
        [LineType.VEPOConsignmentInventoryLine]: "mserp_linenumber",
        [LineType.VERFQAmendment]: "mserp_rfqamendmentnumber",
        [LineType.VEPurchaseOrderResponseArchivedOriginalOrderLine]: "mserp_linenumber",
        [LineType.VEPurchaseOrderConfirmationLine]: "mserp_linenumber"
    }
};

export const FieldNames = {
    HeaderStatus: {
        [LineType.VEPurchaseOrderResponseHeader]: "mserp_responsestate",
        [LineType.VERFQReplyHeader]: "mserp_replyprogressstatus",
        [LineType.VEPendingVendorInvoiceHeader]: "mserp_vendorinvoicereviewstatus",
        [LineType.VEPOConsignmentInventoryHeader]: "mserp_purchaseorderstatus",
        [LineType.VEPurchaseOrderConfirmationHeader]: "mserp_purchaseorderstatus"
    },
    HeaderStatusEdited: {
        [LineType.VEPurchaseOrderResponseHeader]: 200000003,
        [LineType.VERFQReplyHeader]: 200000001,
        [LineType.VEPendingVendorInvoiceHeader]: 200000000,
        [LineType.VEPOConsignmentInventoryHeader]: 200000001
    }
};

export const enum PCFCommunicationType {
    Communication = "vrm:pcfCommunication",
    RefreshAttachments = "vrm:pcfRefreshAttachments",
    RefreshLines = "vrm:pcfRefreshLines"
}

export interface PCFCommunication {
    id: string;
    lineNumberFieldName: string;
    name: number;
    logicalName: string;
    attachmentEntityName: string;
    attachmentLineLookupName: string;
    alternateLineId?: string;
}

export interface RecordsWebApi {
    records: any[];
    hasMoreRecords: boolean;
    pagingCookie: string | undefined;
    pageNumber: number;
}

export type EntityDTO = {
    logicalName: string;
    schemaName: string;
    objectTypeCode?: number;
    displayName?: string;
    displayCollectionName?: string;
    primaryNameAttribute?: string;
    primaryIdAttribute?: string;
    isActivity?: boolean;
    isCustomEntity?: boolean;
    attributes?: AttributeDTO[];
};
export type AttributeDTO = {
    logicalName: string;
    schemaName: string;
    type?: string;
    displayName?: string;
    displayNameLocalized?: LocalizedDTO;
    isValidForRead?: boolean;
    isValidForCreate?: boolean;
    isValidForUpdate?: boolean;
    isCustomAttribute?: boolean;
    targets?: string[];
};

export type LocalizedDTO = {
    [lcid: string]: string;
};