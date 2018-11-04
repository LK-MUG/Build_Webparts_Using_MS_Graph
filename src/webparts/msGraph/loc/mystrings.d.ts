declare interface IMsGraphWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  ClientModeLabel: string;
  SearchFor: string;
  SearchForValidationErrorMessage: string;
}

declare module 'MsGraphWebPartStrings' {
  const strings: IMsGraphWebPartStrings;
  export = strings;
}
