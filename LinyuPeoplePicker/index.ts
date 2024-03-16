import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { ILinyuPeoplePickerProps, LinyuPeoplePickerComponent } from "./LinyuPeoplePicker";
import * as React from "react";
import * as ReactDOM from "react-dom";

export class LinyuPeoplePicker implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private notifyOutputChanged: () => void;
    private theContainer: HTMLDivElement;
    private props: ILinyuPeoplePickerProps = {
        dataSetGrid: [],
        LinyuLabel: "",
        LinyuDisabled: false,
        LinyuSearchText: "",
        LinyuResult: "",
        getComponentsData: this.getComponentsData.bind(this),
        getSearchText: this.getSearchText.bind(this)
    }

    /**
     * Empty constructor.
     */
    // constructor() {

    // }

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
     */
    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement): void {
        // Add control initialization code
        this.notifyOutputChanged = notifyOutputChanged;
        this.theContainer = container;
    }

    public getComponentsData(newResult: string) {
        if (this.props.LinyuResult !== newResult)
            this.props.LinyuResult = newResult;

        this.notifyOutputChanged();
    }

    public getSearchText(searchText: string) {
        if (this.props.LinyuSearchText !== searchText)
            this.props.LinyuSearchText = searchText;

        this.notifyOutputChanged();
    }

    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void {
        // Add code to update control view
        const results: any[] = [];
        const data = context.parameters.dataSetGrid;
        this.props.LinyuLabel = context.parameters.LinyuLabel.raw!;
        this.props.LinyuResult = context.parameters.LinyuResult.raw!;
        this.props.LinyuDisabled = context.parameters.LinyuDisabled.raw!;



        if (!data.loading) {
            try {
                for (let i = 0; i < data.sortedRecordIds.length; i++) {
                    const recordId = data.sortedRecordIds[i];
                    const cRecord = data.records[recordId];
                    const obj: any = {};
                    data.columns.forEach((item, index) => {
                        switch (item.name) {
                            case "DisplayName":
                                obj["text"] = cRecord.getFormattedValue(item.name);
                                break;
                            case "Mail":
                                obj["email"] = cRecord.getFormattedValue(item.name);
                                break;
                            case "JobTitle":
                                obj["secondaryText"] = cRecord.getFormattedValue(item.name);
                                break;
                            case "Id":
                                obj["key"] = cRecord.getFormattedValue(item.name);
                                break;
                            default:
                                obj[item.name] = cRecord.getFormattedValue(item.name);
                        }
                    })
                    if (obj.ID != "")
                        results.push(obj);

                }
                this.props.dataSetGrid = results;
                console.log(results);
            }
            catch {
                this.props.dataSetGrid = [];
            }
        }

        ReactDOM.render(
            React.createElement(
                LinyuPeoplePickerComponent,
                this.props
            ),
            this.theContainer
        );
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
     */
    public getOutputs(): IOutputs {
        return {
            LinyuResult: this.props.LinyuResult,
            LinyuSearchText: this.props.LinyuSearchText
        };
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void {
        // Add code to cleanup control if necessary
        ReactDOM.unmountComponentAtNode(this.theContainer);
    }
}
