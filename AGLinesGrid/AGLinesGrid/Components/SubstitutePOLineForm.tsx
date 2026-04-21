/**
 * @file POLineForm.tsx
 * @description This file contains the form dialog that is used
 * for susbstitution of a Line
 * Upon submitting the form, a new record will be created (Substitute) and the current one
 * (substituted) will be updated
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
    Input,
    makeStyles,
    Field,
    InputProps
} from "@fluentui/react-components";
import { DatePicker } from "@fluentui/react-datepicker-compat";

const useStyles = makeStyles({
    content: {
        display: "flex",
        flexDirection: "column",
        rowGap: "10px"
    }
});

export interface SubstitutePOLineFormProps {
    isOpen: boolean;
    onClose: () => void;
    onSubmit: (data: any) => void;
    getString: (s: string) => string;
    attributesMetadata: any[];
}

export const SubstitutePOLineForm = (props: SubstitutePOLineFormProps) => {
    const styles = useStyles();
    const [formData, setFormData] = React.useState<any>({
        msdyn_name: "",
        msdyn_externalitemnumber: "",
        msdyn_confirmedpurchasequantity: "",
        msdyn_confirmedreceiptdate: new Date()
    });

    const handleFormSubmit = (event: React.FormEvent) => {
        event.preventDefault();
        props.onSubmit(formData);
        resetFormData();
    };

    const handleInputChange: InputProps["onChange"] = (event, data) => {
        const fieldName = event.currentTarget.name;
        setFormData((prev: any) => ({
            ...prev,
            [fieldName]: fieldName === "msdyn_confirmedpurchasequantity" && data.value ? parseInt(data.value) : data.value
        }));
    };

    const handleDateChange = (date: Date | null) => {
        setFormData((prev: any) => ({
            ...prev,
            msdyn_confirmedreceiptdate: date
        }));
    };

    const resetFormData = () => {
        setFormData({
            msdyn_name: "",
            msdyn_externalitemnumber: "",
            msdyn_confirmedpurchasequantity: "",
            msdyn_confirmedreceiptdate: new Date()
        });
    };

    return (
        <Dialog modalType="non-modal" open={props.isOpen} onOpenChange={(event, data) => !data.open && props.onClose()}>
            <DialogSurface aria-describedby={undefined}>
                <form onSubmit={handleFormSubmit}>
                    <DialogBody>
                        <DialogTitle>{props.getString("SubstituteLineTitle")}</DialogTitle>
                        <DialogContent className={styles.content}>
                            <Field label={props.attributesMetadata.find(m => m.LogicalName === "msdyn_name")?.DisplayName?.UserLocalizedLabel?.Label}>
                                <Input
                                    required
                                    value={formData.msdyn_name}
                                    onChange={handleInputChange}
                                    name="msdyn_name"
                                />
                            </Field>

                            <Field label={props.attributesMetadata.find(m => m.LogicalName === "msdyn_confirmedreceiptdate")?.DisplayName?.UserLocalizedLabel?.Label}>
                                <DatePicker
                                    name="msdyn_confirmedreceiptdate"
                                    value={formData.msdyn_confirmedreceiptdate}
                                    onSelectDate={handleDateChange}
                                    placeholder={props.getString("SelectDatePlaceholder")}
                                />
                            </Field>

                            <Field label={props.attributesMetadata.find(m => m.LogicalName === "msdyn_confirmedpurchasequantity")?.DisplayName?.UserLocalizedLabel?.Label}>
                                <Input
                                    required
                                    type="number"
                                    value={formData.msdyn_confirmedpurchasequantity}
                                    name="msdyn_confirmedpurchasequantity"
                                    onChange={handleInputChange}
                                />
                            </Field>

                            <Field label={props.attributesMetadata.find(m => m.LogicalName === "msdyn_externalitemnumber")?.DisplayName?.UserLocalizedLabel?.Label}>
                                <Input
                                    required
                                    value={formData.msdyn_externalitemnumber}
                                    onChange={handleInputChange}
                                    name="msdyn_externalitemnumber"
                                />
                            </Field>
                        </DialogContent>
                        <DialogActions>
                            <Button type="submit" appearance="primary">
                            {props.getString("Submit")}
                            </Button>
                            <DialogTrigger disableButtonEnhancement>
                                <Button appearance="secondary">{props.getString("Close")}</Button>
                            </DialogTrigger>
                        </DialogActions>
                    </DialogBody>
                </form>
            </DialogSurface>
        </Dialog>
    );
};
