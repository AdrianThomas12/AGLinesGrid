import { CustomCellRendererProps } from "ag-grid-react";
import {
    GridProps,
    LineDoneOptions,
    LineStatusOptions,
    LineType,
    PendingVendorLineActionButtons,
    POLineActionButtons,
    RFQLineActionButtons,
    VirtualEntityActions
} from "./types";

export function getRowActions(
    entityName: string,
    params?: {
        handleDeleteRow: (id: string) => Promise<void>;
        executeServerAction?: (gridRowActionName: string, id: string, actionName: string, data?: any) => Promise<any>;
        getString: (str: string) => string;
    }
) {
    const entityActions: Record<
        string,
        Array<{ displayName: string; logicalName: string; onClick: (props: CustomCellRendererProps, data?: any) => Promise<void> | void }>
    > = {
        [LineType.VEPurchaseOrderResponseLine]: [
            {
                displayName: params.getString(POLineActionButtons.RejectLine),
                logicalName: POLineActionButtons.RejectLine,
                onClick: async (props: CustomCellRendererProps) => {
                    const id = props.data[entityName + "id"];
                    await params?.executeServerAction?.(
                        POLineActionButtons.RejectLine,
                        id,
                        VirtualEntityActions[entityName].Reject
                    );
                    props.node.setSelected(false);
                }
            },
            {
                displayName: params.getString(POLineActionButtons.SplitDelivery),
                logicalName: POLineActionButtons.SplitDelivery,
                onClick: async (props: CustomCellRendererProps, data?: any) => {
                    const id = props.data[entityName + "id"];
                    await params?.executeServerAction?.(
                        POLineActionButtons.SplitDelivery,
                        id,
                        VirtualEntityActions[entityName].SplitDelivery,
                        data
                    );
                    props.node.setSelected(false);
                }
            },
            {
                displayName: params.getString(POLineActionButtons.Substitute),
                logicalName: POLineActionButtons.Substitute,
                onClick: async (props: CustomCellRendererProps) => {
                    const id = props.data[entityName + "id"];
                    await params?.executeServerAction?.(
                        POLineActionButtons.Substitute,
                        id,
                        VirtualEntityActions[entityName].Substitute
                    );
                    props.node.setSelected(false);
                }
            },
            {
                displayName: params.getString(POLineActionButtons.DiscardChanges),
                logicalName: POLineActionButtons.DiscardChanges,
                onClick: async (props: CustomCellRendererProps) => {
                    const id = props.data[entityName + "id"];
                    await params?.executeServerAction?.(
                        POLineActionButtons.DiscardChanges,
                        id,
                        VirtualEntityActions[entityName].Discard
                    );
                    props.node.setSelected(false);
                }
            }
        ],
        [LineType.VERFQReplyLine]: [
            {
                displayName: params.getString(RFQLineActionButtons.AddAlternativeLine),
                logicalName: RFQLineActionButtons.AddAlternativeLine,
                onClick: async (props: CustomCellRendererProps) => {
                    const id = props.data[entityName + "id"];
                    await params?.executeServerAction?.(
                        RFQLineActionButtons.AddAlternativeLine,
                        id,
                        VirtualEntityActions[entityName].AddAlternativeLine
                    );
                    props.node.setSelected(false);
                }
            },
            {
                displayName: params.getString(RFQLineActionButtons.RemoveAlternativeLine),
                logicalName: RFQLineActionButtons.RemoveAlternativeLine,
                onClick: async (props: CustomCellRendererProps) => {
                    const id = props.data[entityName + "id"];
                    await params?.executeServerAction?.(
                        RFQLineActionButtons.RemoveAlternativeLine,
                        id,
                        VirtualEntityActions[entityName].RemoveAlternativeLine
                    );
                    props.node.setSelected(false);
                }
            },
            {
                displayName: params.getString(RFQLineActionButtons.ResetLine),
                logicalName: RFQLineActionButtons.ResetLine,
                onClick: async (props: CustomCellRendererProps) => {
                    const id = props.data[entityName + "id"];
                    await params?.executeServerAction?.(
                        RFQLineActionButtons.ResetLine,
                        id,
                        VirtualEntityActions[entityName].ResetLine
                    );
                    props.node.setSelected(false);
                }
            }
        ],
        [LineType.VEPendingVendorInvoiceLine]: [
            {
                displayName: params.getString(PendingVendorLineActionButtons.DeleteLine),
                logicalName: PendingVendorLineActionButtons.DeleteLine,
                onClick: async (props: CustomCellRendererProps) => {
                    const id = props.data[entityName + "id"];
                    await params.handleDeleteRow(id);
                    props.node.setSelected(false);
                }
            }
        ]
    };

    return entityActions[entityName] || [];
}

export function getPOSpecificConfig(configProps: any, configProps2?: any): GridProps {
    let config: GridProps;
    try {
        config = JSON.parse(configProps);
        let config2 = configProps2 ? JSON.parse(configProps2) : {};

        const mergedConfig: GridProps = {
            ...config,
            ...config2,
            customColumns: [...(config.customColumns || []), ...(config2.customColumns || [])]
        };

        if (typeof mergedConfig.hidePriceFields === "string") {
            mergedConfig.hidePriceFields = (mergedConfig.hidePriceFields as string).split(",").map(s => s.trim());
        }

        mergedConfig.customColumns = mergedConfig.customColumns.map((col: any) => {
            const updatedCol = { ...col };

            if (updatedCol.options === "LineStatusOptions") {
                updatedCol.options = LineStatusOptions;
            } else if (updatedCol.options === "LineDoneOptions") {
                updatedCol.options = LineDoneOptions;
            }

            if (updatedCol.editableOnStatus && !Array.isArray(updatedCol.editableOnStatus)) {
                updatedCol.editableOnStatus = [updatedCol.editableOnStatus];
            }

            return updatedCol;
        });

        return mergedConfig as GridProps;
    } catch (error) {
        console.error("Error parsing configuration", error);
        alert("Error parsing configuration");
        return {
            customColumns: [],
            rowActions: [],
            showRowActions: true
        } as GridProps;
    }
}
