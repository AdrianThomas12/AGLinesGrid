/**
 * @file ConfirmationDialog.tsx
 * @description This file contains the confirmation dialog that is used for the User to confirm an action.
 * This is mostly for discarding changes on the Line
 */

import * as React from "react";
import {
    Dialog,
    DialogTrigger,
    DialogSurface,
    DialogTitle,
    DialogContent,
    DialogBody,
    DialogActions,
    Button,
    Spinner
} from "@fluentui/react-components";

export interface ConfirmationDialogProps {
    isOpen: boolean;
    title: string;
    content: string;
    onConfirm: () => void;
    onClose: () => void;
    getString: (s: string) => string;
    confirmText?: string | undefined;
    cancelText?: string | undefined;
}

export const ConfirmationDialog = (props: ConfirmationDialogProps) => {
    const [loading, setLoading] = React.useState<boolean>(false);

    const handleConfirm = async () => {
        setLoading(true);
        try {
            await props.onConfirm();
        } finally {
            setLoading(false);
            props.onClose();
        }
    };

    return (
        <Dialog open={props.isOpen} onOpenChange={(event, data) => !data.open && props.onClose()}>
            <DialogSurface>
                <DialogBody>
                    <DialogTitle>{props.title}</DialogTitle>
                    <DialogContent>{props.content}</DialogContent>
                    <DialogActions>
                        <Button
                            appearance="primary"
                            onClick={handleConfirm}
                            icon={loading ? <Spinner size="tiny" /> : undefined}
                        >
                            {props.confirmText ? props.confirmText : props.getString("OK")}
                        </Button>
                        <DialogTrigger disableButtonEnhancement>
                            <Button appearance="secondary">
                                {props.cancelText ? props.cancelText : props.getString("Close")}
                            </Button>
                        </DialogTrigger>
                    </DialogActions>
                </DialogBody>
            </DialogSurface>
        </Dialog>
    );
};
