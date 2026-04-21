import * as React from "react";
import { Checkbox, FluentProvider, webLightTheme } from "@fluentui/react-components";
import { NoYes } from "../Scripts/types";
import { normalizeCheckboxValue } from "../Scripts/webApi";

export const CheckboxCellEditor = React.forwardRef((props: any, ref: any) => {
    const [checked, setChecked] = React.useState(props.value === NoYes.Yes);
    const checkedRef = React.useRef(checked);

    React.useImperativeHandle(ref, () => ({
        getValue: () => {
            return checked ? NoYes.Yes : NoYes.No;
        },
        isPopup: () => false
    }));

    React.useEffect(() => {
        setChecked(normalizeCheckboxValue(props.value) === NoYes.Yes ? true : false);
        checkedRef.current = normalizeCheckboxValue(props.value) === NoYes.Yes ? true : false;
    }, [props.value]);

    return (
        <FluentProvider theme={webLightTheme}>
            <div className="ag-cell-checkbox-custom">
                <label style={{ cursor: "pointer " }}>
                    <Checkbox
                        checked={checked}
                        onChange={(_, data) => {
                            setChecked(data.checked === true);
                            props.onValueChange?.(data.checked ? NoYes.Yes : NoYes.No);
                            checkedRef.current = data.checked === true;
                        }}
                        tabIndex={0}
                    />
                </label>
            </div>
        </FluentProvider>
    );
});
