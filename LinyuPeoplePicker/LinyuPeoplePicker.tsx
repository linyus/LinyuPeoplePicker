import * as React from 'react';
import { IPersonaProps } from '@fluentui/react/lib/Persona';
import { IBasePickerSuggestionsProps, NormalPeoplePicker, ValidationState } from '@fluentui/react/lib/Pickers';

export interface ILinyuPeoplePickerProps {
    dataSetGrid: any[];
    LinyuLabel: string;
    LinyuDisabled: boolean;
    LinyuResult: string;
    LinyuSearchText: string;
    getComponentsData: (newResult: string) => void;
    getSearchText: (searchText: string) => void;
}

export interface ILinyuPeoplePickerStates {
    LinyuValue: any[];
    peopleList: IPersonaProps[];
}

const suggestionProps: IBasePickerSuggestionsProps = {
    suggestionsHeaderText: 'Suggested People',
    mostRecentlyUsedHeaderText: 'Suggested Contacts',
    noResultsFoundText: 'No results found',
    loadingText: 'Loading',
    showRemoveButtons: false,
    suggestionsAvailableAlertText: 'People Picker Suggestions available',
    suggestionsContainerAriaLabel: 'Suggested contacts',
};

const emptyString = "string.empty";

export class LinyuPeoplePickerComponent extends React.Component<ILinyuPeoplePickerProps, ILinyuPeoplePickerStates>{

    constructor(props: ILinyuPeoplePickerProps) {
        super(props);
        try {
            this.state = {
                LinyuValue: [],
                peopleList: this.props.dataSetGrid
            };
        } catch {
            console.log("Error");
        }
    }

    public componentDidMount() {
        try {
            if (this.props.LinyuResult.indexOf("email") >= 0)
                this.ppl.state.items = JSON.parse(this.props.LinyuResult);
        } catch {
            console.log("Error");
        }
    }

    public componentWillReceiveProps(newprops: any) {
        this.setState({ peopleList: newprops.dataSetGrid });
    }

    protected ppl: any;

    public render(): JSX.Element {
        return (
            <div className={"formField relative"}>
                <div className="label">{this.props.LinyuLabel === emptyString ? undefined : this.props.LinyuLabel}
                </div>
                <div className={"field"}>
                    <NormalPeoplePicker
                        itemLimit={1}
                        defaultSelectedItems={this.state.LinyuValue.length === 0 ? undefined : this.state.LinyuValue}
                        onResolveSuggestions={this.onFilterChanged}
                        getTextFromItem={(persona: IPersonaProps) => { return persona.text as string }}
                        pickerSuggestionsProps={suggestionProps}
                        className={'ms-PeoplePicker'}
                        onChange={(items: any[] | undefined): void => { this.props.getComponentsData(JSON.stringify(items)); }}
                        key={'normal'}
                        onValidateInput={this.validateInput}
                        onInputChange={this.onInputChange}
                        resolveDelay={1000}
                        disabled={this.props.LinyuDisabled}
                        ref={c => (this.ppl = c)}
                    />
                </div>
            </div>
        );
    }

    public listContainsPersona(persona: IPersonaProps, personas: IPersonaProps[]) {
        console.log("listContainsPersona");
        if (!personas || !personas.length || personas.length === 0) {
            return false;
        }
        return personas.filter(item => item.text === persona.text).length > 0;
    }

    public validateInput(input: string): ValidationState {
        console.log("validateInput");
        if (input.indexOf('@') !== -1) {
            return ValidationState.valid;
        } else if (input.length > 1) {
            return ValidationState.warning;
        } else {
            return ValidationState.invalid;
        }
    }

    public onInputChange = (input: string): string => {
        console.log("onInputChange");
        this.props.getSearchText(input);
        const outlookRegEx = /<.*>/g;
        const emailAddress = outlookRegEx.exec(input);

        if (emailAddress && emailAddress[0]) {
            return emailAddress[0].substring(1, emailAddress[0].length - 1);
        }
        return input;
    }

    public onFilterChanged = (filterText: string, currentPersonas: IPersonaProps[] | undefined, limitResults?: number): IPersonaProps[] | Promise<IPersonaProps[]> => {
        console.log("onFilterChanged");
        if (filterText) {
            let filteredPersonas: IPersonaProps[] = this.filterPersonasByText(filterText);
            if (currentPersonas != undefined)
                filteredPersonas = filteredPersonas.filter(persona => !this.listContainsPersona(persona, currentPersonas));

            return filteredPersonas;
        } else {
            return [];
        }
    };

    public filterPersonasByText = (filterText: string): IPersonaProps[] => {
        return this.state.peopleList.filter(item => this.doesTextStartWith(item.text as string, filterText));
    };

    public doesTextStartWith(text: string, filterText: string): boolean {
        return text.toLowerCase().indexOf(filterText.toLowerCase()) >= 0;
    }
}