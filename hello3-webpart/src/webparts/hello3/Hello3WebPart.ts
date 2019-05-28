import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';

import { escape } from '@microsoft/sp-lodash-subset';

import styles from './Hello3WebPart.module.scss';
import * as strings from 'Hello3WebPartStrings';
import  MockHttpClient  from './mockHttpClient';

export interface IHello3WebPartProps {
  description: string;
  sometest: string;
  test1: boolean;
  test2: string;
  test3: boolean;
}

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  Id: string;
}

export default class Hello3WebPart extends BaseClientSideWebPart<IHello3WebPartProps> {

  public render(): void {
    
    var n= MockHttpClient.get();

    this.domElement.innerHTML = `
      <div class="${ styles.hello3 }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
              <p class="${ styles.description }">${escape(this.properties.description)}</p>
              <p class="${ styles.description }">${this.properties.sometest}</p>
              <p class="${ styles.description }">${this.context.pageContext.web.title}</p>c
              <p class="${ styles.description }">${this.getMockListData()}</p>
            </div>
          </div>
        </div>
      </div>`;
  }


  private getMockListData() : Promise<ISPLists> {

    return MockHttpClient.get()
      .then(this.onGetSuccess, this.onGetFailed) as Promise<ISPLists>;


  }

  private onGetSuccess(data : ISPList[])  {

    var l: ISPLists;
    l = {value : null}
    return l;
  }

  
  private onGetFailed(error)  {
    return error;
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
                }),
                PropertyPaneTextField('sometest', {
                  label: 'multiline text',
                  multiline : true
                }),
                PropertyPaneCheckbox('test1', {
                  text: 'checkbox'
                  
                }),
                PropertyPaneDropdown('test2', {
                  label: 'dropdown',
                  options: [
                    { key: '1', text: 'One' }
                  ]}),
                PropertyPaneToggle('test3', {
                  label: 'toggle',
                  onText: 'Yes',
                  offText: 'No'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
