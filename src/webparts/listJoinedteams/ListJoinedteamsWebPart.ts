import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import listJoinedTeamsService, {ListJoinedTeamsService} from './services/ListJoinedTeamsService';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'ListJoinedteamsWebPartStrings';
import ListJoinedteams from './components/ListJoinedteams';
import { IListJoinedteamsProps } from './components/IListJoinedteamsProps';

export interface IListJoinedteamsWebPartProps {
  description: string;
  showIcons: boolean;
  showDummyTeam: boolean;
}

export default class ListJoinedteamsWebPart extends BaseClientSideWebPart<IListJoinedteamsWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IListJoinedteamsProps> = React.createElement(
      ListJoinedteams,
      {
        description: this.properties.description,
        showIcons: this.properties.showIcons,
        showDummyTeam: this.properties.showDummyTeam
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  // protected get dataVersion(): Version {
  //   return Version.parse('1.0');
  // }

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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneToggle('showIcons', {
                  label: strings.ShowIconsFieldLabel
                }),
                PropertyPaneToggle('showDummyTeam', {
                  label: strings.ShowDummyTeamLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }

  protected onInit():Promise<void>{  
    return super.onInit().then(() => {  
      listJoinedTeamsService.setUp(this.context);  
    });  
  }  

}
