import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'VacationListManagerWebPartStrings';
import VacationListManager from './components/VacationListManager';
import { IVacationListManagerProps } from './components/IVacationListManagerProps';

export interface IVacationListManagerWebPartProps {
  description: string;
  url: string;
  context: any;
}

export default class VacationListManagerWebPart extends BaseClientSideWebPart<IVacationListManagerWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IVacationListManagerProps> = React.createElement(
      VacationListManager,
      {
        description: this.properties.description,
        url: this.context.pageContext.web.absoluteUrl,
        context: this.context,
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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
