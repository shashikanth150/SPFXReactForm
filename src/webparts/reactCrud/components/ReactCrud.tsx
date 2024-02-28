import * as React from "react";
import styles from "./ReactCrud.module.scss";
import { IReactCrudProps } from "./IReactCrudProps";

import { ISoftwareListItem } from "./ISoftwareListItem";
import { IReactCrudState } from "./IReactCrudState";

import {
  ISPHttpClientOptions,
  SPHttpClient,
  SPHttpClientResponse,
} from "@microsoft/sp-http";

import {
  TextField,
  PrimaryButton,
  DetailsList,
  DetailsListLayoutMode,
  CheckboxVisibility,
  SelectionMode,
  Dropdown,
  IDropdown,
  IDropdownOption,
  ITextFieldStyles,
  IDropdownStyles,
  Selection,
} from "office-ui-fabric-react";
import * as strings from "ReactCrudWebPartStrings";

let _softwareListColumns = [
  {
    key: "ID",
    name: "ID",
    fieldName: "ID",
    minWidth: 50,
    maxWidth: 100,
    isResizable: true,
  },
  {
    key: "Title",
    name: "Title",
    fieldName: "Title",
    minWidth: 50,
    maxWidth: 100,
    isResizable: true,
  },
  {
    key: "SoftwareName",
    name: "SoftwareName",
    fieldName: "SoftwareName",
    minWidth: 50,
    maxWidth: 100,
    isResizable: true,
  },
  {
    key: "SoftwareVendor",
    name: "SoftwareVendor",
    fieldName: "SoftwareVendor",
    minWidth: 50,
    maxWidth: 100,
    isResizable: true,
  },
  {
    key: "SoftwareVersion",
    name: "SoftwareVersion",
    fieldName: "SoftwareVersion",
    minWidth: 50,
    maxWidth: 100,
    isResizable: true,
  },
  {
    key: "SoftwareDescription",
    name: "SoftwareDescription",
    fieldName: "SoftwareDescription",
    minWidth: 50,
    maxWidth: 150,
    isResizable: true,
  },
];

const textFieldStyles: Partial<ITextFieldStyles> = {
  fieldGroup: { width: 300 },
};
const narrowDropdownStyles: Partial<IDropdownStyles> = {
  dropdown: { width: 300 },
};

export default class ReactCrud extends React.Component<
  IReactCrudProps,
  IReactCrudState
> {
  private _selection: Selection;

  private _onItemsSelectionChanged = () => {
    const selectedItems = this._selection.getSelection();

    this.setState({
      SoftwareListItem: selectedItems[0] as ISoftwareListItem,
    });
    console.log(selectedItems);
  };

  constructor(props: IReactCrudProps, state: IReactCrudState) {
    super(props);
    this.state = {
      status: "Ready",
      SoftwareListItems: [],
      SoftwareListItem: {
        Id: 2,
        Title: "",
        SoftwareName: "",
        SoftwareDescription: "",
        SoftwareVendor: "Select an option",
        SoftwareVersion: "",
      },
    };

    this._selection = new Selection({
      onSelectionChanged: this._onItemsSelectionChanged,
    });
  }

  private _getListItems(): Promise<ISoftwareListItem[]> {
    const url: string =
      this.props.siteUrl +
      "/_api/web/lists/getbytitle('Shashi%20Catalog')/items";
    return this.props.context.spHttpClient
      .get(url, SPHttpClient.configurations.v1)
      .then((response) => {
        return response.json();
      })
      .then((json) => {
        return json.value;
      }) as Promise<ISoftwareListItem[]>;
  }
  public bindDetailsList(message: string): void {
    this._getListItems().then((listItems) => {
      this.setState({ SoftwareListItems: listItems, status: message });
    });
  }

  public componentDidMount(): void {
    this.bindDetailsList("All Records have been loaded Successfully");
  }

  public btnAdd_click = (): void => {
    const url: string =
      this.props.siteUrl +
      "/_api/web/lists/getbytitle('Shashi%20Catalog')/items";
    const spHttpClientOptions: ISPHttpClientOptions = {
      body: JSON.stringify(this.state.SoftwareListItem),
    };

    this.props.context.spHttpClient
      .post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
      .then((response: SPHttpClientResponse) => {
        if (response.status === 201) {
          this.bindDetailsList(
            "Record added and all records were loaded successfully"
          );
        } else {
          let errorMessage: string =
            "An error has occurred: " +
            response.status +
            " - " +
            response.statusText;
          this.setState({ status: errorMessage });
        }
      });
  };

  public btnUpdate_click = (): void => {
    let id: number =
      this.state.SoftwareListItem && this.state.SoftwareListItem.Id
        ? this.state.SoftwareListItem.Id
        : 0; //this.state.SoftwareListItem.Id;

    if (id !== undefined || null) {
      const url: string =
        this.props.siteUrl +
        "/_api/web/lists/getbytitle('Shashi%20Catalog')/items(" +
        id +
        ")";
      const headers: any = {
        "X-HTTP-Method": "MERGE",
        "IF-MATCH": "*",
      };

      const spHttpClientOptions: ISPHttpClientOptions = {
        headers: headers,
        body: JSON.stringify(this.state.SoftwareListItem),
      };

      this.props.context.spHttpClient
        .post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
        .then((response: SPHttpClientResponse) => {
          if (response.status === 204) {
            this.bindDetailsList(
              "Record Updated and all records were loaded successfully"
            );
          } else {
            let errorMessage: string =
              "An error has occurred: " +
              response.status +
              " - " +
              response.statusText;
            this.setState({ status: errorMessage });
          }
        });
    } else {
      return;
    }
  };

  public btnDelete_click = (): void => {
    let id: number =
      this.state.SoftwareListItem && this.state.SoftwareListItem.Id
        ? this.state.SoftwareListItem.Id
        : 0; //this.state.SoftwareListItem.Id;
    if (id !== undefined || null) {
      const url: string =
        this.props.siteUrl +
        "/_api/web/lists/getbytitle('Shashi%20Catalog')/items(" +
        id +
        ")";

      const headers: any = { "X-HTTP-Method": "DELETE", "IF-MATCH": "*" };

      const spHttpClientOptions: ISPHttpClientOptions = {
        headers: headers,
      };

      this.props.context.spHttpClient
        .post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
        .then((response: SPHttpClientResponse) => {
          if (response.status === 204) {
            alert("record got deleted successfully....");
            this.bindDetailsList(
              "Record deleted and All Records were loaded Successfully"
            );
          } else {
            let errormessage: string =
              "An error has occured i.e.  " +
              response.status +
              " - " +
              response.statusText;
            this.setState({ status: errormessage });
          }
        });
    } else {
      return;
    }
  };
  public render(): React.ReactElement<IReactCrudProps> {
    const dropdownRef = React.createRef<IDropdown>();

    return (
      <div className={styles.reactCrud}>
        <TextField
          label={strings.lblID}
          required={false}
          // value={(this.state.SoftwareListItem.Id).toString()}
          value={
            this.state.SoftwareListItem && this.state.SoftwareListItem.Id
              ? this.state.SoftwareListItem.Id.toString()
              : ""
          }
          styles={textFieldStyles}
          onChange={(e) => {
            const newValue = (
              e.target as HTMLInputElement | HTMLTextAreaElement
            ).value;
            this.setState((prevState) => ({
              SoftwareListItem: {
                ...prevState.SoftwareListItem,
                Id: Number(newValue),
              },
            }));
          }}
        />
        <TextField
          label={strings.lblSoftwareTitle}
          required={true}
          value={
            this.state.SoftwareListItem && this.state.SoftwareListItem.Title
              ? this.state.SoftwareListItem.Title
              : ""
          } //{ (this.state.SoftwareListItem.Title)}
          styles={textFieldStyles}
          onChange={(e) => {
            const newValue = (
              e.target as HTMLInputElement | HTMLTextAreaElement
            ).value;
            this.setState((prevState) => ({
              SoftwareListItem: {
                ...prevState.SoftwareListItem,
                Title: newValue,
              },
            }));
          }}
        />
        <TextField
          label={strings.lblSoftwareName}
          required={true}
          value={
            this.state.SoftwareListItem &&
            this.state.SoftwareListItem.SoftwareName
              ? this.state.SoftwareListItem.SoftwareName
              : ""
          } //{ (this.state.SoftwareListItem.SoftwareName)}
          styles={textFieldStyles}
          onChange={(e) => {
            const newValue = (
              e.target as HTMLInputElement | HTMLTextAreaElement
            ).value;
            this.setState((prevState) => ({
              SoftwareListItem: {
                ...prevState.SoftwareListItem,
                SoftwareName: newValue,
              },
            }));
          }}
        />
        <TextField
          label={strings.lblSoftwareDescription}
          required={true}
          value={
            this.state.SoftwareListItem &&
            this.state.SoftwareListItem.SoftwareDescription
              ? this.state.SoftwareListItem.SoftwareDescription
              : ""
          } //{ (this.state.SoftwareListItem.SoftwareDescription)}
          styles={textFieldStyles}
          onChange={(e) => {
            const newValue = (
              e.target as HTMLInputElement | HTMLTextAreaElement
            ).value;
            this.setState((prevState) => ({
              SoftwareListItem: {
                ...prevState.SoftwareListItem,
                SoftwareDescription: newValue,
              },
            }));
          }}
        />
        <TextField
          label={strings.lblSoftwareVersion}
          required={true}
          value={
            this.state.SoftwareListItem &&
            this.state.SoftwareListItem.SoftwareVersion
              ? this.state.SoftwareListItem.SoftwareVersion
              : ""
          } //{ (this.state.SoftwareListItem.SoftwareVersion)}
          styles={textFieldStyles}
          onChange={(e) => {
            const newValue = (
              e.target as HTMLInputElement | HTMLTextAreaElement
            ).value;
            this.setState((prevState) => ({
              SoftwareListItem: {
                ...prevState.SoftwareListItem,
                SoftwareVersion: newValue,
              },
            }));
          }}
        />
        <Dropdown
          componentRef={dropdownRef}
          placeholder="Select an option"
          label={strings.lblSoftwareVendor}
          options={[
            { key: "Microsoft", text: "Microsoft" },
            { key: "Sun", text: "Sun" },
            { key: "Oracle", text: "Oracle" },
            { key: "Google", text: "Google" },
          ]}
          defaultSelectedKey={
            this.state.SoftwareListItem &&
            this.state.SoftwareListItem.SoftwareVendor
              ? this.state.SoftwareListItem.SoftwareVendor
              : ""
          } //{this.state.SoftwareListItem.SoftwareVendor}
          required={true}
          styles={narrowDropdownStyles}
          onChange={(
            e: React.FormEvent<HTMLDivElement>,
            option?: IDropdownOption,
            index?: number
          ): void => {
            if (option) {
              const newValue = option.key as string;
              this.setState((prevState) => ({
                SoftwareListItem: {
                  ...prevState.SoftwareListItem,
                  SoftwareVendor: newValue,
                },
              }));
            }
          }}
        />

        <p className={styles.title}>
          <PrimaryButton text="Add" title="Add" onClick={this.btnAdd_click} />
          &nbsp;
          <PrimaryButton text="Update" onClick={this.btnUpdate_click} />
          &nbsp;
          <PrimaryButton text="Delete" onClick={this.btnDelete_click} />
        </p>

        <div id="divStatus">{this.state.status}</div>

        <div>
          <DetailsList
            items={this.state.SoftwareListItems}
            columns={_softwareListColumns}
            setKey="Id"
            checkboxVisibility={CheckboxVisibility.always}
            selectionMode={SelectionMode.single}
            layoutMode={DetailsListLayoutMode.fixedColumns}
            compact={true}
            selection={this._selection}
          />
        </div>
      </div>
    );
  }
}
