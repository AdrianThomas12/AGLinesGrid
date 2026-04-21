import * as React from "react";
import { Button, Tooltip } from "@fluentui/react-components";
import { AppsListDetailRegular, AppsListDetailFilled } from "@fluentui/react-icons";

export interface DetailsTooltipProps {
  getString: (s: string) => string;
}

export const DetailsTooltip = (props: DetailsTooltipProps) => {
    const [visible, setVisible] = React.useState(false);
    const [enabled, setEnabled] = React.useState(false);
    return (
        <Tooltip
            content="The checkbox controls whether the tooltip can show on hover or focus"
            relationship="description"
            visible={visible && enabled}
            onVisibleChange={(_ev, data) => setVisible(data.visible)}
            positioning="below"
        >
            <Button
                onClick={() => setEnabled(!enabled)}
                icon={enabled ? <AppsListDetailFilled /> : <AppsListDetailRegular />}
            />
        </Tooltip>
    );
};
