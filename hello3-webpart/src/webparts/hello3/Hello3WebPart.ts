import { Version, Environment, EnvironmentType } from '@microsoft/sp-core-library';
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
import {SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

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
              <p class="${ styles.description }">${this.context.pageContext.web.title}</p>
              
            </div>
          </div>
          <div id="spListContainer"></div>
        </div>
      </div>`;

      this._renderListAsync();
  }


  private _getMockListData() : Promise<ISPLists> {

    return MockHttpClient.get()
      .then(this.onGetSuccess, this.onGetFailed) as Promise<ISPLists>;


  }

  private onGetSuccess(data : ISPList[])  {

    var listsData: ISPLists = {value : data}
    return listsData;
  }

  
  private onGetFailed(error)  {
    return error;
  }

  private _getListData(): Promise<ISPLists> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists?$filter=Hidden eq false`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  private _renderList(items: ISPList[]): void {
    let html: string = '';
    items.forEach((item: ISPList) => {
      html += `
    <ul class="${styles.list}">
      <li class="${styles.listItem}">
        <span class="ms-font-l">${item.Title}</span>
      </li>
    </ul>`;
    });

    const listContainer: Element = this.domElement.querySelector('#spListContainer');
    listContainer.innerHTML = html;
  }

  private _renderListAsync(): void {

    if (Environment.type === EnvironmentType.Local) {
      this._getMockListData().then((response) => {
        this._renderList(response.value);
      });
    }
    else if (Environment.type == EnvironmentType.SharePoint || 
              Environment.type == EnvironmentType.ClassicSharePoint) {
      this._getListData()
        .then((response) => {
          this._renderList(response.value);
        });
    }
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
