/**
 * @file App.tsx
 * @description The root component of the application that renders the grid within a Fluent UI provider.
 */

import * as React from "react";
import { Grid } from "./Grid";
import { FluentProvider, IdPrefixProvider, webLightTheme } from "@fluentui/react-components";

/**
 * The main application component that wraps the Grid component inside a Fluent UI provider.
 *
 * @param {any} props - The properties passed to the application.
 * @returns {JSX.Element} The rendered App component.
 */
export const App = (props: any) => {
    const controlRef = React.useRef<HTMLDivElement>(null);
    return (
        <IdPrefixProvider value={props.controlId}>
            <div id="eg-root" ref={controlRef}>
                <FluentProvider theme={webLightTheme}>
                    <Grid {...props} controlRef={controlRef} />
                </FluentProvider>
            </div>
        </IdPrefixProvider>
    );
};
