import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './HelloWorld4WebPart.module.scss';
import * as strings from 'HelloWorld4WebPartStrings';

import { PropertyPaneToggle, IPropertyPaneToggleProps } from '@microsoft/sp-property-pane';

const toggleProps: IPropertyPaneToggleProps = { 
  label: "Toggl1", 
  onText: "Yes", 
  offText: "No" 
};
  

export interface IHelloWorld4WebPartProps {
  description: string,
  settingsToggle: boolean;

}

export default class HelloWorld4WebPart extends BaseClientSideWebPart<IHelloWorld4WebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.helloWorld4 }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
              <p class="${ styles.description }">${escape(this.properties.description)}</p>
              <a href="https://aka.ms/spfx" class="${ styles.button }">
                <span class="${ styles.label }">Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
      <div>${this.properties.settingsToggle}</div>
      `;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
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
            }
          ]
        },
        {
          header : {
              description: "Header2-mydesc2"
          },
          groups : [
            {
              groupName: "grp1",
              groupFields : [
                PropertyPaneToggle("settingsToggle", toggleProps)
              ]

            },
            {
              groupName: "grp2",
              groupFields : [
                
              ]

            },
          ]
          
        }
      ]
    };
  }
}
