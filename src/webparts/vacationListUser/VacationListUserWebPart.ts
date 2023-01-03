import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'VacationListUserWebPartStrings';
import VacationListUser from './components/VacationListUser';
import { IVacationListUserProps } from './components/IVacationListUserProps';

export interface IVacationListUserWebPartProps {
  description: string;
}

export default class VacationListUserWebPart extends BaseClientSideWebPart<IVacationListUserWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IVacationListUserProps> = React.createElement(
      VacationListUser,
      {
        description: this.properties.description,
        context: this.context,
        url: this.context.pageContext.web.absoluteUrl
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
