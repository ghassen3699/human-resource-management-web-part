import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'VacationRequestWebPartStrings';
import VacationRequest from './components/VacationRequest';
import { IVacationRequestProps } from './components/IVacationRequestProps';
import { sp } from '@pnp/sp';

export interface IVacationRequestWebPartProps {
  description: string;
  context: any;
  url: any;
  LanguageSelected: number;
}

export default class VacationRequestWebPart extends BaseClientSideWebPart<IVacationRequestWebPartProps> {

  protected onInit(): Promise<void> {
    return super.onInit().then(_ => {
      // other init code may be present
      sp.setup({
        spfxContext: this.context
      });
    });
  }
  
  public render(): void {
    const element: React.ReactElement<IVacationRequestProps> = React.createElement(
      VacationRequest,
      {
        description: this.properties.description,
        context: this.context,
        url: this.context.pageContext.web.absoluteUrl,
        LanguageSelected: this.properties.LanguageSelected,
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




  constructor() {
    super();
    Object.defineProperty(this, "dataVersion", {
      get() {
        return Version.parse('1.0');
      }
    }),
    Object.defineProperty(this, "disableReactivePropertyChanges", {
      get(){
        return true;
      }
    })
  };
  

  protected onAfterPropertyPaneChangesApplied(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
    this.render();
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            },
            // config of webpart name
            {
              groupName: 'Language Config',
              groupFields: [
                PropertyPaneDropdown('LanguageSelected',{
                  label: 'Choose your language',
                  selectedKey: 1,
                  disabled: false,
                  options:[
                    {
                      key:1,
                      text:"Frensh"
                    },
                    {
                      key:2,
                      text:"Arabic"
                    },
                    {
                      key:3,
                      text:"English"
                    },
                  ]

                })
              ]
            }
          ]
        }
      ]
    };
  }
}
