declare interface IHelloWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
}

declare module 'HelloWebPartStrings' {
  const strings: IHelloWebPartStrings;
  export = strings;
}
