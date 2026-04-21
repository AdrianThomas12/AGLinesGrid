import {
    FluentProvider,
    webLightTheme,
    makeStyles,
    Button,
    Input,
    PopoverSurface,
    Popover,
    PositioningImperativeRef,
    Skeleton,
    SkeletonItem
} from "@fluentui/react-components";
import { DismissFilled, DocumentDatabaseRegular, SearchRegular } from "@fluentui/react-icons";
import * as React from "react";
import { DropdownTable } from "./DropdownTable";
import * as lodash from "lodash";
import { RecordsWebApi } from "../../Scripts/types";
import { getFetchXml, getLookupValues } from "../../Scripts/webApi";

/**
 * Lookup JSX element properties interface.
 */
export interface LookupProps {
    /** Controls context */
    webApi: ComponentFramework.WebApi;
    /** The control display name */
    columnDisplayName: string;
    /** The currently selected value to display in the input field.  */
    selectedValue: ComponentFramework.LookupValue;
    /** The columns to be displayed in the dropdown table. */
    columns: { key: string; label: string }[];
    /** A flag indicating whether the input should be styled as a portal input or a regular input. */
    isPortal: boolean;
    /** Is the value being retrieved */
    isValueLoading: boolean;
    /** Is the control disabled */
    isDisabled: boolean;
    /** A callback function to handle when a value is selected from the dropdown. */
    onChange: (value: any) => void;
    /** A function that fetches values for the dropdown based on a search query.  */
    getValues: (searchText: string, currentPage: number) => Promise<RecordsWebApi>;
    /** Get Localized control strings. */
    getString(id: string): string;
    /** The field to display on the lookup control */
    displayField: string;
    /** The table image as svg */
    entityIconSvg: string | null;
    /** Configuration with columns */
    configuration: {
        lookupEntityType: string;
        displayField: string;
        secondaryDisplayField: string;
        filterRecordId: string;
        filterFieldRelationColumn: string;
        saveTextFieldName: string;
    };
    /** The vendor id */
    vendorId: string | null;
}

/**
 * Styles for the control using Fluent UI's makeStyles
 * Defines the layout and appearance of the control components
 */
const useStyles = makeStyles({
    root: {
        display: "flex",
        height: "32px",
    },
    inputWrapper: {
        width: "100%"
    },
    input: {
        flex: 1,
        width: "100%",
        paddingLeft: "4px !important",
        backgroundColor: "#FFFF"
    },
    searchButton: {
        width: "20px",
        height: "20px",
        marginRight: "-10px"
    },
    popoverSurface: {
        padding: "2px"
    },
    tagContainer: {
        display: "flex",
        alignItems: "center",
        backgroundColor: "rgb(235, 243, 252)",
        borderRadius: "4px",
        padding: "2px 0px 0px 0px",
        marginRight: "4px",
        cursor: "pointer",
        overflow: "hidden",
        maxWidth: "250px",
        ":hover": {
            backgroundColor: "rgb(207, 228, 250)"
        }
    },
    tagText: {
        fontSize: "14px",
        marginRight: "4px",
        overflow: "hidden",
        textOverflow: "ellipsis",
        whiteSpace: "nowrap",
        color: "rgb(17, 94, 163)",
        textDecoration: "underline",
        maxWidth: "180px"
    },
    tagRemoveButton: {
        background: "none",
        border: "none",
        fontSize: "14px",
        cursor: "pointer",
        color: "rgb(17, 94, 163)",
        padding: "0",
        width: "20px",
        maxWidth: "20px",
        minWidth: "20px"
    },
    entityImage: {
        marginRight: "5px"
    },
    iconWrapper: {
        width: "24px",
        display: "flex",
        alignItems: "center",
        justifyContent: "center"
    }
});

/**
 * React component that renders the custom lookup display
 */
export const Lookup = (props: LookupProps) => {
    const styles = useStyles();

    const [recordsWebApi, setRecordsWebApi] = React.useState<RecordsWebApi>({
        records: [],
        hasMoreRecords: false,
        pagingCookie: undefined,
        pageNumber: 1
    });
    const [lookup, setLookup] = React.useState<ComponentFramework.LookupValue>(null);
    const [isValueLoading, setValueLoading] = React.useState<boolean>(props.isValueLoading);
    const [loading, setLoading] = React.useState(false);
    const [isDropdownOpen, setIsDropdownOpen] = React.useState(false);
    const [inputWidth, setInputWidth] = React.useState(0);
    const [searchText, setSearchText] = React.useState("");
    const [debouncedSearchText, setDebouncedSearchText] = React.useState("");
    const [inputText, setInputText] = React.useState("");
    const [inputValue, setInputValue] = React.useState("");
    const [inputFocused, setInputFocused] = React.useState(false);

    const inputRef = React.useRef<HTMLInputElement>(null);
    const inputWrapperRef = React.useRef<HTMLDivElement>(null);
    const positioningRef = React.useRef<PositioningImperativeRef>(null);

    const dropdownWidth = React.useMemo(() => Math.max(inputWidth, 490), [inputWidth]);
    const recordsScrollContainerRef = React.useRef<HTMLDivElement>(null);

    const handleInputFocus = async () => {
        setInputFocused(true);

        if (!isDropdownOpen) {
            setLoading(true);
            if (!lookup?.id) {
                setInputText("");
            }
            setSearchText("");
            setDebouncedSearchText("");
            setIsDropdownOpen(true);

            const newValues = await props.getValues(searchText, recordsWebApi.pageNumber);

            setRecordsWebApi(newValues);
            setLoading(false);
        }
    };

    const handleInputBlur = () => {
        setInputFocused(false);
    };

    const handleInputChange = (e: any, data: any) => {
        setInputText(data.value);
        setSearchText(data.value);
    };

    const handleClearValue = async () => {
        setInputFocused(true);
        setLookup(null);
        setInputText("");
        setSearchText("");
        setDebouncedSearchText("");
        props.onChange(null);
    };

    const handleOpenChange = (e: any, data: any) => {
        if (e.target === inputRef.current) {
            inputRef.current?.focus();
            // Ignore events that are triggered by the input to avoid re-opening the popover
            return;
        }

        if (inputFocused && !data.open) {
            return;
        }

        setIsDropdownOpen(data.open);

        if (!data.open) {
            setInputText(inputValue);
        }
    };

    const handleValueClick = (value: any) => {
        props.onChange(value);

        setInputValue(value.name || props.getString("NoName"));
        setInputText(value.name || props.getString("NoName"));
        setLookup(value);
        setIsDropdownOpen(false);
    };

    const updateInputWidth = React.useCallback(() => {
        if (inputRef.current) {
            setInputWidth(inputRef.current.offsetWidth + 28); // fill in for the modified padding of PopoverSurface
        }
    }, []);

    React.useEffect(() => {
        updateInputWidth();

        window.addEventListener("resize", updateInputWidth);

        return () => {
            window.removeEventListener("resize", updateInputWidth);
        };
    }, [updateInputWidth]);

    React.useEffect(() => {
        if (inputRef.current) {
            positioningRef.current?.setTarget(inputRef.current);
        }
    }, [inputRef, positioningRef]);

    React.useEffect(() => {
        const handler = setTimeout(() => {
            setDebouncedSearchText(searchText);
        }, 300);

        return () => {
            clearTimeout(handler);
        };
    }, [searchText]);

    React.useEffect(() => {
        const updateValues = async () => {
            const newValues = await props.getValues(searchText, recordsWebApi.pageNumber);

            setRecordsWebApi(newValues);
            setLoading(false);
        };

        setLoading(true);
        updateValues();

        if (isDropdownOpen) {
            inputRef.current?.focus();
        }
    }, [debouncedSearchText]);

    React.useEffect(() => {
        const handleWindowBlur = () => {
            setIsDropdownOpen(false);
        };

        window.addEventListener("blur", handleWindowBlur);

        return () => {
            window.removeEventListener("blur", handleWindowBlur);
        };
    }, []);

    React.useEffect(() => {
        setInputText(
            props.selectedValue && props.selectedValue?.id ? props.selectedValue.name || props.getString("NoName") : ""
        );
        setInputValue(
            props.selectedValue && props.selectedValue?.id ? props.selectedValue.name || props.getString("NoName") : ""
        );
    }, [props.selectedValue]);

    React.useEffect(() => {
        if (!props.selectedValue?.id || !props.displayField) return;

        setLookup(props.selectedValue);
        setInputText(props.selectedValue?.name || props.getString("NoName"));
        setInputValue(props.selectedValue?.name || props.getString("NoName"));
    }, [props.selectedValue]);

    /**
     * Handles infinite scrolling for loading additional records in the attachments and notes sections.
     * This function is throttled to limit the frequency of scroll event handling and avoid performance issues.
     *
     * When the user scrolls to the bottom of the container, it checks if more records are available
     *
     * @param {React.UIEvent<HTMLDivElement>} e - Event of the HTML div element.
     * @param {ListItemType} type - Type of the attachment record.
     */
    const handleScroll = React.useCallback(
        lodash.debounce(async (e: React.UIEvent<HTMLDivElement> | Event) => {
            const eventTarget = e.target as HTMLDivElement;

            if (eventTarget.scrollHeight === 0) return;

            if (Math.abs(eventTarget.scrollHeight - eventTarget.scrollTop - eventTarget.clientHeight) < 5) {
                if (!loading) {
                    setLoading(true);
                    const previousScrollHeight = eventTarget.scrollHeight;
                    const fetchXml = getFetchXml(
                        props.configuration.lookupEntityType,
                        props.configuration.displayField,
                        props.configuration.secondaryDisplayField,
                        props.configuration.filterRecordId,
                        props.configuration.filterFieldRelationColumn,
                        searchText,
                        recordsWebApi.pageNumber,
                        true,
                        props.vendorId,
                        recordsWebApi.pagingCookie
                    );

                    const newRecordsWebApi = await getLookupValues(
                        fetchXml,
                        props.configuration.lookupEntityType,
                        props.configuration.displayField,
                        props.configuration.secondaryDisplayField,
                        recordsWebApi.pageNumber,
                        props.webApi
                    );
                    setRecordsWebApi((prev: any) => ({
                        records: [...prev.records, ...newRecordsWebApi.records],
                        pagingCookie: newRecordsWebApi.pagingCookie,
                        hasMoreRecords: newRecordsWebApi.hasMoreRecords,
                        pageNumber: newRecordsWebApi.pageNumber
                    }));
                    setLoading(false);
                    requestAnimationFrame(() => {
                        eventTarget.scrollTop = eventTarget.scrollHeight - previousScrollHeight;
                    });
                }
            }
        }, 500),
        [recordsWebApi]
    );

    React.useEffect(() => {
        const recordsContainer = recordsScrollContainerRef.current;

        if (!recordsContainer) return;

        const handleRecordsScroll = (e: Event) => handleScroll(e as React.UIEvent<HTMLDivElement> | Event);

        if (recordsContainer) {
            recordsContainer.addEventListener("scroll", handleRecordsScroll);
        }

        return () => {
            if (recordsContainer) {
                recordsContainer.removeEventListener("scroll", handleRecordsScroll);
            }
        };
    }, [handleScroll]);

    return (
        <div style={{ width: "95%"}}>
            <FluentProvider theme={webLightTheme}>
                {isValueLoading ? (
                    <Skeleton aria-label="Loading Content">
                        <SkeletonItem size={32} />
                    </Skeleton>
                ) : (
                    <div className={styles.root}>
                        <Popover
                            open={isDropdownOpen}
                            positioning={{
                                target: inputWrapperRef.current,
                                position: "below",
                                align: "bottom",
                                offset: 2
                            }}
                            onOpenChange={handleOpenChange}
                        >
                            <div ref={inputWrapperRef} className={styles.inputWrapper}>
                                <Input
                                    placeholder={
                                        lookup?.id
                                            ? ""
                                            : props
                                                  .getString("LookForTableName")
                                                  .replace("{0}", props.columnDisplayName)
                                    }
                                    ref={inputRef}
                                    className={styles.input}
                                    appearance={props.isPortal ? "outline" : "filled-darker"}
                                    readOnly={!!lookup?.id}
                                    contentAfter={
                                        <Button
                                            appearance="transparent"
                                            className={styles.searchButton}
                                            icon={<SearchRegular />}
                                            onClick={handleInputFocus}
                                        />
                                    }
                                    contentBefore={
                                        lookup?.id ? (
                                            <div className={styles.tagContainer}>
                                                <div className={styles.iconWrapper}>
                                                    {props.entityIconSvg ? (
                                                        <div
                                                            className="entity-image-svg-wrapper"
                                                            dangerouslySetInnerHTML={{ __html: props.entityIconSvg }}
                                                        />
                                                    ) : (
                                                        <DocumentDatabaseRegular
                                                            color="rgb(17, 94, 163)"
                                                            style={{ fontSize: 14, color: "rgb(17, 94, 163)" }}
                                                        />
                                                    )}
                                                </div>

                                                <span className={styles.tagText}>{inputText}</span>
                                                <Button
                                                    appearance="transparent"
                                                    className={styles.tagRemoveButton}
                                                    icon={
                                                        <DismissFilled
                                                            color="rgb(17, 94, 163)"
                                                            style={{ fontSize: 14, color: "rgb(17, 94, 163)" }}
                                                        />
                                                    }
                                                    onClick={handleClearValue}
                                                />
                                            </div>
                                        ) : null
                                    }
                                    value={lookup?.id ? "" : inputText}
                                    onChange={handleInputChange}
                                    onFocus={handleInputFocus}
                                    onBlur={handleInputBlur}
                                    disabled={props.isDisabled}
                                />
                            </div>
                            <PopoverSurface
                                className={`${styles.popoverSurface} custom-lookup-popover`}
                                onFocus={() => inputRef.current?.focus()}
                                style={{ width: dropdownWidth }}
                                data-lookup-popup
                            >
                                <DropdownTable
                                    columns={props.columns}
                                    items={recordsWebApi.records || []}
                                    loading={loading}
                                    width={dropdownWidth}
                                    onRowClick={row => handleValueClick(row)}
                                    getString={props.getString}
                                    scrollContainerRef={recordsScrollContainerRef}
                                    handleScroll={handleScroll}
                                />
                            </PopoverSurface>
                        </Popover>
                    </div>
                )}
            </FluentProvider>
        </div>
    );
};
