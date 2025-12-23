declare interface IDynamicFormBuilderWebPartStrings {
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
  UnknownEnvironment: string;
}

declare module 'DynamicFormBuilderWebPartStrings' {
  const strings: IDynamicFormBuilderWebPartStrings;
  export = strings;
}

declare module "*.module.scss" {
  const classes: { [key: string]: string };
  export default classes;
}