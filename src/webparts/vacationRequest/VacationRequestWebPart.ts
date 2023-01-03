import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'VacationRequestWebPartStrings';
import VacationRequest from './components/VacationRequest';
import { IVacationRequestProps } from './components/IVacationRequestProps';

export interface IVacationRequestWebPartProps {
  description: string;
  context: any;
  url: any;
}

export default class VacationRequestWebPart extends BaseClientSideWebPart<IVacationRequestWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IVacationRequestProps> = React.createElement(
      VacationRequest,
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
