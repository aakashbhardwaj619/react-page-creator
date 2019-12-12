import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IPropertyPaneDropdownOption } from '@microsoft/sp-webpart-base';

export interface IPageCreatorProps {
  selectedSites: string[];
  context: WebPartContext;
  followedSites: IPropertyPaneDropdownOption[];
  buttonText: string;
  panelHeading: string;
  featuredSitesHeading: string;
  buttonAlignment: string;
}
