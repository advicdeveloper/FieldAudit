import * as React from "react";
import { createRoot, Root } from "react-dom/client";
import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { HistoryTooltip, IHistoryTooltipProps } from "./HistoryTooltip";

export class FieldAuditTrail implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private _container: HTMLDivElement;
    private _context: ComponentFramework.Context<IInputs>;
    private _notifyOutputChanged: () => void;
    private _root: Root;

    /**
     * Empty constructor.
     */
    constructor() {
        // Empty
    }

    /**
     * Used to initialize the control instance.
     */
    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        this._context = context;
        this._notifyOutputChanged = notifyOutputChanged;
        this._container = container;
        this._root = createRoot(this._container);
    }

    /**
     * Called when any value in the property bag has changed.
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this._context = context;
        const props: IHistoryTooltipProps = {
            context: context,
            value: context.parameters.value.raw || "",
            fieldName: context.parameters.value.attributes?.LogicalName || "unknown",
            onChange: (newValue: string | undefined) => {
                context.parameters.value.raw = newValue || null;
                this._notifyOutputChanged();
            }
        };

        this._root.render(React.createElement(HistoryTooltip, props));
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     */
    public getOutputs(): IOutputs {
        return {
            value: this._context.parameters.value.raw || undefined
        };
    }

    /**
     * Called when the control is to be removed from the DOM tree.
     */
    public destroy(): void {
        this._root.unmount();
    }
}
