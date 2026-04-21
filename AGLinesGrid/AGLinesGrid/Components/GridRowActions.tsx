/**
 * @file GridRowActions.tsx
 * @description Component that renders action buttons for each row in the grid.
 */

import * as React from "react";
import {
    FluentProvider,
    Menu,
    MenuButton,
    MenuItem,
    MenuList,
    MenuPopover,
    MenuTrigger,
    Spinner,
    webLightTheme
} from "@fluentui/react-components";
import { MoreHorizontalRegular } from "@fluentui/react-icons";
import { CustomCellRendererProps } from "ag-grid-react";
import {
    AttachmentEntityName,
    AttachmentFieldName,
    GridRowActionProps,
    LineSplitDeliveryAction,
    LineStatus,
    LineType,
    NoYes,
    PCFCommunication,
    PCFCommunicationType,
    PendingVendorLineActionButtons,
    POLineActionButtons,
    RFQLineActionButtons,
    RFQLineType
} from "../Scripts/types";
import "./App.css";
import { ConfirmationDialog } from "./ConfirmationDialog";
import { SplitDeliveryDrawer } from "./SplitDeliveryDrawer";
import { normalizeCheckboxValue } from "../Scripts/webApi";

// Allowed actions per Line Status
const poAllowedActions: Record<LineStatus, string[]> = {
    [LineStatus.Rejected]: [POLineActionButtons.DiscardChanges],
    [LineStatus.Substituted]: [POLineActionButtons.DiscardChanges],
    [LineStatus.Substitute]: [], // No actions allowed
    [LineStatus.SplitIntoSchedule]: [POLineActionButtons.SplitDelivery, POLineActionButtons.DiscardChanges],
    [LineStatus.ScheduleLine]: [], // No actions allowed
    [LineStatus.Accepted]: [
        POLineActionButtons.RejectLine,
        POLineActionButtons.Substitute,
        POLineActionButtons.SplitDelivery
    ],
    [LineStatus.AcceptedWithChanges]: [POLineActionButtons.DiscardChanges],
    [LineStatus.SplitFromSubstitute]: [] // No actions allowed
};

export const GridRowActions = (
    props: CustomCellRendererProps & {
        actions: GridRowActionProps[];
        refreshGridData: () => void;
        entityName: string;
        poHeaderId: string;
        poHeaderSchemaName: string;
        poHeaderEntityName: string;
        lineStatus: number;
        getString: (s: string) => string;
        attributesMetadata: any[];
        isDiscardEnabled?: boolean;
        linesData?: LineSplitDeliveryAction;
        controlId: string;
        webAPI: ComponentFramework.WebApi;
        debugMode?: boolean;
    }
) => {
    const [selectedRow, setSelectedRow] = React.useState<any>(null);
    const [loading, setLoading] = React.useState<boolean>(false);
    const [confirmDialogOpen, setConfirmDialogOpen] = React.useState<boolean>(false);
    const [selectedAction, setSelectedAction] = React.useState<GridRowActionProps | null>(null);
    const [drawerOpened, setDrawerOpen] = React.useState<boolean>(false);
    const availablePOActions = poAllowedActions[props.lineStatus as LineStatus] || [];
    const [confirmationMessage, setConfirmationMessage] = React.useState<string>("");

    const handle_Action_OnClick = React.useCallback(
        async (action: GridRowActionProps) => {
            if (action.logicalName === POLineActionButtons.DiscardChanges) {
                setSelectedAction(action);
                setConfirmationMessage(props.getString("DiscardChangesConfirmation"));
                setConfirmDialogOpen(true);
            } else if (action.logicalName === RFQLineActionButtons.ResetLine) {
                setSelectedAction(action);
                setConfirmationMessage(props.getString("ResetLineConfirmation"));
                setConfirmDialogOpen(true);
            } else if (action.logicalName === POLineActionButtons.SplitDelivery) {
                setSelectedAction(action);
                setSelectedRow(props.data);
                setDrawerOpen(true);
            } else if (action.logicalName === PendingVendorLineActionButtons.DeleteLine) {
                const isCreate = !!props.data?.isCreate;
                if (isCreate) {
                    props.api.applyTransaction({ remove: [props.data] });
                } else {
                    setLoading(true);
                    await action.onClick(props);
                    setLoading(false);
                    props.refreshGridData();
                }
            } else {
                try {
                    setLoading(true);
                    await action.onClick(props);
                    setLoading(false);
                    props.refreshGridData();
                } catch (error) {
                    setLoading(false);
                    alert(error);
                }

                refreshAttachmentsCustomEvent();
            }
        },
        [props]
    );

    const handleDiscardAction = async () => {
        try {
            setLoading(true);
            setConfirmDialogOpen(false);
            await selectedAction.onClick(props);
            setLoading(false);
            props.refreshGridData();
            refreshAttachmentsCustomEvent();
        } catch (error) {
            setLoading(false);
            alert(error);
        }
    };

    const handleSplitDeliveryAction = async (data: LineSplitDeliveryAction) => {
        try {
            setDrawerOpen(false);
            setLoading(true);
            await selectedAction.onClick(props, data);
            setSelectedRow(null);
            setLoading(false);
            props.refreshGridData();
            refreshAttachmentsCustomEvent();
        } catch (error) {
            setLoading(false);
            alert(error);
        }
    };

    /**
     * Dispatches a custom event to refresh Attachments
     */
    const refreshAttachmentsCustomEvent = () => {
        var customEvent = new CustomEvent<PCFCommunication>(
            `${PCFCommunicationType.RefreshAttachments}:${props.controlId}`,
            {
                detail: {
                    id: props.data[props.entityName + "id"],
                    lineNumberFieldName: AttachmentFieldName.LineNumber[props.entityName],
                    name:
                        typeof props.data.mserp_linenumber === "string"
                            ? parseFloat(props.data.mserp_linenumber)
                            : props.data.mserp_linenumber,
                    logicalName: props.entityName,
                    attachmentEntityName: AttachmentEntityName[props.entityName],
                    attachmentLineLookupName: AttachmentFieldName.Lookup[props.entityName]
                }
            }
        );
        window.dispatchEvent(customEvent);
    };

    return (
        <div className="action-menu-cell">
            <FluentProvider theme={webLightTheme}>
                <Menu>
                    <MenuTrigger disableButtonEnhancement>
                        <MenuButton
                            icon={loading ? <Spinner size="tiny" /> : <MoreHorizontalRegular />}
                            appearance="subtle"
                        />
                    </MenuTrigger>

                    <MenuPopover>
                        <MenuList>
                            {props.actions.map((a, i) => {
                                if (
                                    !availablePOActions.includes(a.logicalName) &&
                                    props.entityName === LineType.VEPurchaseOrderResponseLine
                                )
                                    return null; // Hide actions not allowed

                                if (
                                    props.entityName === LineType.VERFQReplyLine &&
                                    props.data.mserp_linetype &&
                                    props.data.mserp_linetype !== RFQLineType.Category &&
                                    (a.logicalName === RFQLineActionButtons.AddAlternativeLine ||
                                        a.logicalName === RFQLineActionButtons.RemoveAlternativeLine)
                                ) {
                                    return null;
                                }

                                if (
                                    (a.logicalName === RFQLineActionButtons.RemoveAlternativeLine &&
                                        normalizeCheckboxValue(props.data.mserp_isalternateproduct) === NoYes.No) ||
                                    (a.logicalName === RFQLineActionButtons.AddAlternativeLine &&
                                        (normalizeCheckboxValue(props.data.mserp_allowalternates) === NoYes.No ||
                                            normalizeCheckboxValue(props.data.mserp_isalternateproduct) === NoYes.Yes))
                                ) {
                                    return null;
                                }

                                return (
                                    <MenuItem
                                        disabled={a.disabled}
                                        key={`gridrowaction-menu-item${i}`}
                                        onClick={() => handle_Action_OnClick(a)}
                                    >
                                        {a.displayName}
                                    </MenuItem>
                                );
                            })}
                        </MenuList>
                    </MenuPopover>
                </Menu>

                <ConfirmationDialog
                    isOpen={confirmDialogOpen}
                    title={props.getString("ConfirmationDialogTitle")}
                    content={confirmationMessage}
                    onConfirm={handleDiscardAction}
                    onClose={() => setConfirmDialogOpen(false)}
                    getString={props.getString}
                />

                <SplitDeliveryDrawer
                    open={drawerOpened}
                    headerId={props.poHeaderId}
                    existing={selectedRow?.mserp_linestate === "Split into schedule"}
                    setOpen={setDrawerOpen}
                    title={props.getString("SplitDelivery")}
                    data={selectedRow}
                    handleAction={handleSplitDeliveryAction}
                    getString={props.getString}
                    webAPI={props.webAPI}
                    debugMode={props.debugMode}
                />
            </FluentProvider>
        </div>
    );
};
