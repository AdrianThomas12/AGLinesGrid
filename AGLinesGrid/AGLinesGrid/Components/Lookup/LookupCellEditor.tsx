import * as React from "react";
import { Lookup } from "./Lookup";

export const LookupCellEditor = React.forwardRef((props: any, ref: any) => {
    const [value, setValue] = React.useState(props.value ?? null);

    React.useImperativeHandle(ref, () => ({
        getValue: () => {
            return value ? { ...value } : value;
        },
        isPopup: () => true
    }));

    React.useEffect(() => {
        setValue(props.value ?? null);
    }, [props.value]);

    return (
        <Lookup
            {...props}
            selectedValue={value}
            onChange={(val: any) => {
                setValue(val);
                props.onValueChange(val);
            }}
        />
    );
});