import * as React from "react";
import { Popover, PopoverSurface, PopoverTrigger, Button, Text, Divider } from "@fluentui/react-components";
import { Dismiss24Regular, AppsListDetailFilled } from "@fluentui/react-icons";

export const DetailsPopover = ({
    data,
    fields,
    additionalDetailsText
}: {
    data: any;
    additionalDetailsText: string;
    fields: {
        fieldName: string;
        label: string;
    }[];
}) => {
    const [open, setOpen] = React.useState(false);

    return (
        <Popover open={open} onOpenChange={(_, d) => setOpen(d.open)} positioning="below">
            <PopoverTrigger disableButtonEnhancement>
                <Button
                    icon={<AppsListDetailFilled />}
                    appearance="subtle"
                    size="small"
                    onClick={() => setOpen(true)}
                />
            </PopoverTrigger>

            <PopoverSurface>
                <div style={{ padding: "0.75rem", minWidth: 200, maxWidth: 250, position: "relative" }}>
                    <Text weight="semibold" block style={{ marginBottom: "0.25rem" }}>
                        {additionalDetailsText}
                    </Text>
                    <Button
                        icon={<Dismiss24Regular />}
                        appearance="transparent"
                        size="small"
                        onClick={() => setOpen(false)}
                        style={{ position: "absolute", top: 6, right: 6 }}
                    />
                    <Divider />
                    <div style={{ marginTop: "0.75rem" }}>
                        {fields.map((f, i) => (
                            <div key={i} style={{ marginBottom: "0.5rem" }}>
                                <Text size={300} block>
                                    {f.label}
                                </Text>
                                <Text size={400} weight="semibold">
                                    {data?.[f.fieldName] ?? "—"}
                                </Text>
                            </div>
                        ))}
                    </div>
                </div>
            </PopoverSurface>
        </Popover>
    );
};
