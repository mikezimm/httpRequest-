import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './HttpRequestWebPart.module.scss';
import * as strings from 'HttpRequestWebPartStrings';

import { functionUpdateGroup, UpdateGroup } from "../../services/httpRequest";

export interface IHttpRequestWebPartProps {
  description: string;
}

const siteUrl = 'https://mcclickster.sharepoint.com/sites/PivotNotInstalled'; //Must be a valid SharePoint Url with the Id and groups below
const siteID = 'fbeb30c8-3c2a-492d-bacf-1bd6686c9d35';  //Must be Id of Site, not Web
const ownerGroupID = '84';  //Must be a valid SharePoint Group ID
const targetGroupId = '103';  //Must be a valid SharePoint Group ID

export default class HttpRequestWebPart extends BaseClientSideWebPart<IHttpRequestWebPartProps> {
  
  private groupService: UpdateGroup;
  private groupResultElement: HTMLElement;

  protected onInit(): Promise<void>{
    this.groupService = new UpdateGroup( this.context.httpClient );
    return Promise.resolve();
  }

  public render(): void {

    let result = functionUpdateGroup( this.context.httpClient, siteUrl, siteID, targetGroupId, ownerGroupID );

    let result2 = JSON.stringify(result);

    if ( !this.renderedOnce ) {
      this.domElement.innerHTML = `
      <div class="${ styles.httpRequest }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Welcome to SharePoint!</span>

              <h2>This is calling the class</h2>
              <div style="padding: 20px 0px" class="groupResultClass">Add result here</div>

              <h2>This is calling the function</h2>
              <div style="padding: 20px 0px" class="groupResultClass2">${result2}</div>

            </div>
          </div>
        </div>
      </div>`;
    }

    this.groupResultElement = this.domElement.getElementsByClassName('groupResultClass')[0] as HTMLElement;

    this.assignGroupOwner();

  }

  public assignGroupOwner() : void {
    this.groupService.updateOwner( siteUrl, siteID, targetGroupId, ownerGroupID )
      .then(( results: any ) => {
        this.groupResultElement.innerHTML = JSON.stringify( results );
      })
      .catch(( error ) => {
        alert('Had a problem doing update' + error.message );
      });

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
        }
      ]
    };
  }
}
