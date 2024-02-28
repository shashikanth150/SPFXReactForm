declare interface IReactCrudWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  lblID: string;
  lblSoftwareTitle: string;
  lblSoftwareName: string;
  lblSoftwareDescription: string;
  lblSoftwareVendor: string;
  lblSoftwareVersion: string;
}

declare module 'ReactCrudWebPartStrings' {
  const strings: IReactCrudWebPartStrings;
  export = strings;
}
