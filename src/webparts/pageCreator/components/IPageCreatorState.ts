import { IChoiceGroupOption } from 'office-ui-fabric-react/lib/ChoiceGroup';

export interface IPageCreatorState {
    showPanel: boolean;
    featuredSiteProperties: any[];
    followedSiteProperties: any[];
    followedSitesLinkText: string;
    featuredSitesLinkText: string;
    showTemplates: boolean;
    templateOptions: IChoiceGroupOption[];
    loading: boolean;
    selectedSiteUrl: string;
    selectedTemplateId: string;
    pageType: string;
}