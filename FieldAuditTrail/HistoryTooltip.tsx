import * as React from "react";
import { IInputs } from "./generated/ManifestTypes";

import { TooltipHost } from "@fluentui/react/lib/Tooltip";
import { Icon } from "@fluentui/react/lib/Icon";
import { Stack, IStackTokens, IStackStyles } from "@fluentui/react/lib/Stack";
import { TextField } from "@fluentui/react/lib/TextField";
import { ActivityItem } from "@fluentui/react/lib/ActivityItem";
import { Link } from "@fluentui/react/lib/Link";
import { Text } from "@fluentui/react/lib/Text";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import { Separator } from "@fluentui/react/lib/Separator";
import { Callout, DirectionalHint } from "@fluentui/react/lib/Callout";
import { IconButton } from "@fluentui/react/lib/Button";

export interface IHistoryTooltipProps {
    context: ComponentFramework.Context<IInputs>;
    value: string | null;
    fieldName: string;
    onChange: (newValue: string | undefined) => void;
}

interface IAuditRecord {
    createdon: string;
    userid: string;
    operation: string;
    action: string;
    auditId: string;
    oldValue?: string;
    newValue?: string;
}

interface IOptionMetadata {
    Options?: {
        Value: string | number;
        Label: string;
    }[];
    Type?: string;
}

const stackTokens: IStackTokens = { childrenGap: 10 };
const tooltipStackStyles: IStackStyles = {
    root: {
        padding: "15px",
        maxWidth: "350px",
        background: "rgba(255, 255, 255, 0.95)",
        backdropFilter: "blur(10px)",
        borderRadius: "8px",
        boxShadow: "0 4px 15px rgba(0,0,0,0.1)",
        border: "1px solid #edebe9"
    }
};

export const HistoryTooltip: React.FunctionComponent<IHistoryTooltipProps> = (props) => {
    const { context, value, onChange, fieldName } = props;

    const [history, setHistory] = React.useState<IAuditRecord[]>([]);
    const [loading, setLoading] = React.useState<boolean>(false);
    const [error, setError] = React.useState<string | null>(null);
    const [hasFetched, setHasFetched] = React.useState(false);
    const [isCalloutVisible, setIsCalloutVisible] = React.useState(false);
    const iconRef = React.useRef<HTMLDivElement>(null);

    const fetchHistory = React.useCallback(async (force = false) => {
        if (loading) return;
        if (hasFetched && !force) return;
        setHasFetched(true);
        setLoading(true);
        setError(null);

        const rawEntityId = (context.mode as unknown as { contextInfo: { entityId: string } }).contextInfo?.entityId;
        if (!rawEntityId) {
            setLoading(false);
            return;
        }

        const entityId = rawEntityId.replace(/\{|\}/g, "");

        try {
            const result = await context.webAPI.retrieveMultipleRecords(
                "audit",
                `?$filter=_objectid_value eq ${entityId} &$orderby=createdon desc &$top=30 &$select=createdon,operation,action,changedata,auditid &$expand=userid($select=fullname)`
            );

            const filteredRecords: IAuditRecord[] = [];

            for (const e of result.entities) {
                if (filteredRecords.length >= 5) break;

                let oldValue: string | undefined;
                let newValue: string | undefined;
                let fieldChanged = false;

                const isCreate = e["operation"] === 1;
                const opLabel = isCreate ? "Created" : "Updated";
                const rawChangeData = e["changedata"];

                if (rawChangeData) {
                    try {
                        if (rawChangeData.trim().startsWith("{")) {
                            const json = JSON.parse(rawChangeData);
                            if (json.changedAttributes && Array.isArray(json.changedAttributes)) {
                                for (const attr of json.changedAttributes) {
                                    if (attr.logicalName?.toLowerCase() === fieldName.toLowerCase()) {
                                        fieldChanged = true;
                                        oldValue = attr.oldValue === null ? "(empty)" : String(attr.oldValue);
                                        newValue = attr.newValue === null ? "(empty)" : String(attr.newValue);
                                        break;
                                    }
                                }
                            }
                        } else {
                            const parser = new DOMParser();
                            const xmlDoc = parser.parseFromString(rawChangeData, "text/xml");
                            const attributes = Array.from(xmlDoc.getElementsByTagName("attribute"));
                            for (const attr of attributes) {
                                if (attr.getAttribute("name")?.toLowerCase() === fieldName.toLowerCase()) {
                                    fieldChanged = true;
                                    const oldNode = attr.getElementsByTagName("oldValue")[0];
                                    const newNode = attr.getElementsByTagName("newValue")[0];
                                    oldValue = oldNode?.textContent || "(empty)";
                                    newValue = newNode?.textContent || "(empty)";
                                    break;
                                }
                            }
                        }
                    } catch (err) {
                        console.error("Error parsing changedata", err);
                    }
                }

                if (!fieldChanged && isCreate) {
                    fieldChanged = true;
                    oldValue = "(initial)";
                    newValue = value || "(empty)";
                }

                if (fieldChanged) {
                    let oldDisplay = oldValue;
                    let newDisplay = newValue;

                    const attrMetadata = context.parameters.value.attributes;
                    if (attrMetadata) {
                        const meta = attrMetadata as unknown as IOptionMetadata;
                        if (meta.Type === "OptionSet" || meta.Type === "TwoOptions") {
                            const options = meta.Options;
                            if (options && Array.isArray(options)) {
                                const oldOpt = options.find(o => String(o.Value) === oldValue);
                                const newOpt = options.find(o => String(o.Value) === newValue);
                                if (oldOpt) oldDisplay = oldOpt.Label;
                                if (newOpt) newDisplay = newOpt.Label;
                            }
                        }
                    }

                    filteredRecords.push({
                        createdon: e["createdon"],
                        userid: e["userid"]?.fullname || "System",
                        operation: opLabel,
                        action: e["action@OData.Community.Display.V1.FormattedValue"] || "",
                        auditId: e["auditid"],
                        oldValue: oldDisplay,
                        newValue: newDisplay
                    });
                }
            }

            setHistory(filteredRecords);
        } catch (err: unknown) {
            if (err instanceof Error) {
                setError(err.message);
            } else {
                setError("Error fetching history");
            }
        } finally {
            setLoading(false);
        }
    }, [context, fieldName, value, hasFetched]);

    const onRenderContent = () => {
        if (loading) {
            return (
                <Stack horizontalAlign="center" tokens={{ padding: 20 }}>
                    <Spinner size={SpinnerSize.medium} label="Fetching history..." />
                </Stack>
            );
        }

        if (error) {
            return (
                <Stack tokens={{ padding: 10 }}>
                    <Text variant="small" style={{ color: "#d13438" }}>{error}</Text>
                </Stack>
            );
        }

        if (history.length === 0) {
            return (
                <Stack tokens={{ padding: 15 }} horizontalAlign="center">
                    <Text variant="medium" style={{ fontWeight: 600 }}>No history found</Text>
                    <Text variant="smallPlus" style={{ color: "#605e5c", textAlign: "center" }}>
                        Auditing for "{fieldName}" might be disabled or no changes were recorded yet.
                    </Text>
                </Stack>
            );
        }

        return (
            <Stack tokens={stackTokens} styles={tooltipStackStyles}>
                <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                    <Text variant="large" style={{ fontWeight: 600, color: "#323130" }}>Field History</Text>
                    <IconButton
                        iconProps={{ iconName: "Cancel" }}
                        onClick={() => setIsCalloutVisible(false)}
                        styles={{ root: { color: "#323130", height: 24, width: 24 } }}
                    />
                </Stack>
                <Separator />
                <Stack tokens={{ childrenGap: 15 }}>
                    {history.map((h) => (
                        <ActivityItem
                            key={h.auditId}
                            activityDescription={[
                                <Link key={1} style={{ fontWeight: 600 }}>{h.userid}</Link>,
                                <span key={2}> {h.operation.toLowerCase()} this field</span>
                            ]}
                            activityIcon={<Icon iconName="Edit" style={{ color: "#0078d4" }} />}
                            timeStamp={new Date(h.createdon).toLocaleString()}
                            comments={
                                <Stack tokens={{ childrenGap: 5 }} style={{ marginTop: 4 }}>
                                    <Stack horizontal tokens={{ childrenGap: 5 }}>
                                        <Text variant="small" style={{ fontWeight: 600, color: "#a4262c" }}>Old Value:</Text>
                                        <Text variant="small" style={{ fontStyle: "italic", userSelect: "text" }}>{h.oldValue}</Text>
                                    </Stack>
                                    <Stack horizontal tokens={{ childrenGap: 5 }}>
                                        <Text variant="small" style={{ fontWeight: 600, color: "#107c10" }}>New Value:</Text>
                                        <Text variant="small" style={{ fontStyle: "italic", userSelect: "text" }}>{h.newValue}</Text>
                                    </Stack>
                                </Stack>
                            }
                        />
                    ))}
                </Stack>
                <Separator />
                <Text variant="tiny" style={{ color: "#a19f9d", textAlign: "right" }}>
                    Showing last {history.length} changes
                </Text>
            </Stack>
        );
    };

    return (
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} style={{ width: "100%" }}>
            <TextField
                value={value || ""}
                onChange={(e, v) => onChange(v)}
                borderless
                autoComplete="off"
                styles={{
                    root: { flexGrow: 1 },
                    fieldGroup: {
                        background: "transparent",
                        borderBottom: "1px solid #deecf9",
                        selectors: {
                            ":after": { borderBottomColor: "#0078d4" }
                        }
                    },
                    field: { fontSize: "14px", color: "#323130" }
                }}
            />
            <div
                ref={iconRef}
                onClick={() => {
                    const newState = !isCalloutVisible;
                    setIsCalloutVisible(newState);
                    if (newState) fetchHistory(true);
                }}
                style={{
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "center",
                    width: "28px",
                    height: "28px",
                    borderRadius: "50%",
                    cursor: "pointer",
                    transition: "background 0.2s",
                    background: isCalloutVisible ? "#deecf9" : "#f3f2f1",
                    color: "#0078d4"
                }}
                onMouseOver={(e) => !isCalloutVisible && (e.currentTarget.style.background = "#deecf9")}
                onMouseOut={(e) => !isCalloutVisible && (e.currentTarget.style.background = "#f3f2f1")}
            >
                <Icon iconName="History" styles={{ root: { fontSize: "14px" } }} />
            </div>
            {isCalloutVisible && (
                <Callout
                    target={iconRef}
                    onDismiss={() => setIsCalloutVisible(false)}
                    directionalHint={DirectionalHint.bottomRightEdge}
                    gapSpace={10}
                    beakWidth={10}
                    styles={{
                        beak: { background: "#fff" },
                        beakCurtain: { background: "transparent" },
                        calloutMain: { borderRadius: "8px", border: "none", outline: "none" }
                    }}
                    setInitialFocus
                >
                    {onRenderContent()}
                </Callout>
            )}
        </Stack>
    );
};
