declare interface IPageCreatorWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  TextPropertiesGroupName: string;
  DescriptionFieldLabel: string;
  ButtonTextFieldLabel: string;
  PanelHeadingFieldLabel: string;
  FeaturedSitesTextFieldLabel: string;
  SelectedSitesFieldLabel: string;
  FollowedSitesFieldLabel: string;
  ButtonAlignmentFieldLabel: string;
}

declare module 'PageCreatorWebPartStrings' {
  const strings: IPageCreatorWebPartStrings;
  export = strings;
}
