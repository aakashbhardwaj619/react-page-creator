import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  PropertyPaneToggle,
  PropertyPaneChoiceGroup,
  IPropertyPaneField,
  PropertyPaneCheckbox
} from '@microsoft/sp-webpart-base';

import * as strings from 'PageCreatorWebPartStrings';
import PageCreator from './components/PageCreator';
import { IPageCreatorProps } from './components/IPageCreatorProps';
import { PropertyFieldMultiSelect } from '@pnp/spfx-property-controls/lib/PropertyFieldMultiSelect';
import { SPService } from '../../service/SPService';
//import { PropertyPaneDropdown } from '@microsoft/sp-property-pane';

export interface IPageCreatorWebPartProps {
  selectedSites: string[];
  followedSites: IPropertyPaneDropdownOption[];
  showFollowedSites: boolean;
  buttonText: string;
  panelHeading: string;
  featuredSitesHeading: string;
  buttonAlignment: string;
  searchSite: string;
}

export default class PageCreatorWebPart extends BaseClientSideWebPart<IPageCreatorWebPartProps> {

  private selectedSites: IPropertyPaneDropdownOption[];
  private optionProps: IPropertyPaneField<any>[];

  public render(): void {
    const element: React.ReactElement<IPageCreatorProps> = React.createElement(
      PageCreator,
      {
        selectedSites: this.properties.selectedSites,
        context: this.context,
        followedSites: this.properties.followedSites,
        buttonText: this.properties.buttonText,
        buttonAlignment: this.properties.buttonAlignment,
        panelHeading: this.properties.panelHeading,
        featuredSitesHeading: this.properties.featuredSitesHeading
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    const _ = await super.onInit();

    if (this.properties.showFollowedSites) {
      SPService.GETFOLLOWEDSITES(this.context.msGraphClientFactory).then((followedSites) => {
        console.log('followed sites: ', followedSites);
        this.properties.followedSites = followedSites;
        this.render();
      });
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyPaneConfigurationStart() {
    /* SPService.GETALLSITES(this.context.msGraphClientFactory).then((selectedSites) => {
      this.selectedSites = selectedSites;
      this.context.propertyPane.refresh();
    }); */
    let optionProps: IPropertyPaneField<any>[] = [];
    if (this.properties.selectedSites) {
      this.properties.selectedSites.map((site) => {
        let index = site.indexOf('###');
        let siteUrl = site.substring(0, index);
        let siteTitle = site.substring(index + 3);

        optionProps.push(PropertyPaneCheckbox(siteUrl,
          {
            text: siteTitle,
            checked: true,
            disabled: true
          })
        );
      });
    }
    this.optionProps = optionProps;
    this.context.propertyPane.refresh();
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any) {
    if (propertyPath === 'showFollowedSites') {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      if (newValue === true) {
        SPService.GETFOLLOWEDSITES(this.context.msGraphClientFactory).then((followedSites) => {
          this.properties.followedSites = followedSites;
        });
      } else {
        this.properties.followedSites = [];
      }
      this.render();
      this.context.propertyPane.refresh();
    }
    if (propertyPath === 'searchSite') {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      if (newValue === '') {
        this.selectedSites = [];
        this.context.propertyPane.refresh();
      } else {
        SPService.GETALLSITES(this.context.msGraphClientFactory, newValue).then((selectedSites) => {
          this.selectedSites = selectedSites;
          this.context.propertyPane.refresh();
        });
      }
    }
    if (propertyPath === 'selectedSites') {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      let optionProps: IPropertyPaneField<any>[] = [];
      this.properties.selectedSites.map((site) => {
        let index = site.indexOf('###');
        let siteUrl = site.substring(0, index);
        let siteTitle = site.substring(index + 3);

        optionProps.push(PropertyPaneCheckbox('siteUrl',
          {
            text: siteTitle,
            checked: true,
            disabled: true
          })
        );
      });
      this.optionProps = optionProps;
      this.context.propertyPane.refresh();
    }
    if (propertyPath === 'siteUrl') {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      console.log('newvalue: ', newValue);
      console.log(`old value: ${oldValue}, new value: ${newValue}`);
      this.context.propertyPane.refresh();
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('searchSite', {
                  label: 'Search for site',
                  placeholder: 'Search by Site Title or type * to get all sites'
                }),
                PropertyFieldMultiSelect('selectedSites', {
                  key: 'selectedSites',
                  label: '',
                  options: this.selectedSites,
                  selectedKeys: this.properties.selectedSites ? this.properties.selectedSites : []
                }),
                ...this.optionProps,
                PropertyPaneToggle('showFollowedSites', {
                  label: strings.FollowedSitesFieldLabel,
                  onText: 'Yes',
                  offText: 'No'
                })
              ]
            },
            {
              groupName: strings.TextPropertiesGroupName,
              groupFields: [
                PropertyPaneTextField('buttonText', {
                  label: strings.ButtonTextFieldLabel
                }),
                PropertyPaneChoiceGroup('buttonAlignment', {
                  label: strings.ButtonAlignmentFieldLabel,
                  options: [
                    { key: 'left', text: 'Left', iconProps: { officeFabricIconFontName: 'AlignLeft' } },
                    { key: 'right', text: 'Right', iconProps: { officeFabricIconFontName: 'AlignRight' } }
                  ]
                }),
                PropertyPaneTextField('panelHeading', {
                  label: strings.PanelHeadingFieldLabel
                }),
                PropertyPaneTextField('featuredSitesHeading', {
                  label: strings.FeaturedSitesTextFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
