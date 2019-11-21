import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './WgoBiReportsV2WebPart.module.scss';
import * as strings from 'WgoBiReportsV2WebPartStrings';

import { SPHttpClient, SPHttpClientConfiguration, SPHttpClientResponse, ODataVersion, ISPHttpClientConfiguration } from '@microsoft/sp-http';
import { IODataUser, IODataWeb } from '@microsoft/sp-odata-types'

import * as $ from 'jquery';
import { SPUser } from '@microsoft/sp-page-context';

export interface IWgoBiReportsV2WebPartProps {
  description: string;
}
export interface DefList {
  UserNameId: number;
  IconURL: string;
  Title: string;
  ReportName: string;
  AuthorId: string;
  ReportURL: string;
  ReportNameLU: string;
}
export interface DefLists {
  value: DefList[];
 }

export interface ISPList {
  UserNameId: number;
  IconURL: string;
  Title: string;
  ReportName: string;
  AuthorId: string;
  ReportURL: string;
  ReportNameLU: string;
}
export interface ISPLists {
  value: ISPList[];
 }
 export interface SPUserIDs {
  value: SPUserID[];
}

export interface SPUserID {
  CurrentUser: number;
}

export interface BIReportList {
  Title: string;
  ID: number;
  ReportURL: string;
  IconLink: string;
  PowerBIType: string;
  Description: string;
  HomePage: string;

}
export interface BIReports {
  value: BIReportList[];

}

export default class WgoBiReportsV2WebPart extends BaseClientSideWebPart<IWgoBiReportsV2WebPartProps> {
  private _getdefaultBiReports(): Promise<DefLists> {
    // return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getByTitle('Locations')/items?$filter=Title eq'Eden Prairie'`, SPHttpClient.configurations.v1)
        const spHttpClient: SPHttpClient = this.context.spHttpClient;
        const currentWebUrl: string = this.context.pageContext.web.absoluteUrl;
      
          // Get the Web URL
          return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getByTitle('BI-Reports')/items?$filter=Default eq 'Yes'&$orderby=Order desc`, SPHttpClient.configurations.v1)
          //?$filter=User eq 'Scott Lassiter' 
          .then((response: SPHttpClientResponse) => {
              return response.json();
              
            });
  }
  private _getMyBiReports(): Promise<ISPLists> {
    // return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getByTitle('Locations')/items?$filter=Title eq'Eden Prairie'`, SPHttpClient.configurations.v1)
        const spHttpClient: SPHttpClient = this.context.spHttpClient;
        const currentWebUrl: string = this.context.pageContext.web.absoluteUrl;
      
          // Get the Web URL
          return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getByTitle('BI-Reports')/items?$orderby=Order desc`, SPHttpClient.configurations.v1)
          //?$filter=User eq 'Scott Lassiter' 
          .then((response: SPHttpClientResponse) => {
              return response.json();
              
            });
  }
  // This funcitona adds the ability for us to force reports if the Selections are set as HomePage
  private _getBIReportSelections(): Promise<BIReports> {
    const spHttpClient: SPHttpClient = this.context.spHttpClient;
    const currentWebUrl: string = this.context.pageContext.web.absoluteUrl;

    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getByTitle('BI-Report-Selections')/items`, SPHttpClient.configurations.v1).
    then((response:SPHttpClientResponse) => {
      return response.json();
    });
  }
  private _getUserID(): Promise<SPUserID> {
    const spHttpClient: SPHttpClient = this.context.spHttpClient;
    const currentWebUrl: string = this.context.pageContext.web.absoluteUrl;

   return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/currentuser`, SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse) => {
      return response.json();
    });
    
   
  }

  private _renderListAsync(): void {
    // Gets the data from the list referenced in getLocations
    const spHttpClient: SPHttpClient = this.context.spHttpClient;
    const currentWebUrl: string = this.context.pageContext.web.absoluteUrl;
    let renderUser: number;
    let biReports: any[];
    let defaultReports: any[];
    this._getBIReportSelections().then((response) => {
       
      biReports = response.value;
    });
    //let biReports: string;
    spHttpClient.get(`${currentWebUrl}/_api/web/currentuser`, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
    
    this._getdefaultBiReports().then((response) => {
       
        defaultReports = response.value;
    });
     response.json().then((user: IODataUser) => {
       renderUser = user.Id;
       console.log("User IDEE" + user.Id);
       
       this._getMyBiReports()
       .then((response) => {
         this._renderList(response.value, defaultReports, biReports, renderUser);
       });
     });
   });
  

   
    console.log("render User: " + renderUser)
  }

  private _renderList(items: ISPList[], defaultReports: DefList[],reports:BIReportList[], User): void {
   
    let html: string = '';
     html += `
          <div class="${styles.wgoBiReports}">     
            <div class=${styles.grid}">
              <div class="${styles.row}">
     `
    // Default reports
    defaultReports.forEach((defaultReport: DefList) => {
      var iconDef = defaultReport.IconURL;
       html += `
            <div class="${styles["bi-Main"]}"><a href="${defaultReport.ReportURL}" target="_blank"><div class="ms-md2  ${styles["bi-holder"]}" style='background-image:url("${iconDef}");'><div class='${styles["app-text"]}'>${defaultReport.Title}</div></div></a></div>
           `
               
    }); 
     items.forEach((item: ISPList) => {
       var icon = item.IconURL;
             
       if (item.UserNameId == User)  {
         html += `
             <div class="${styles["bi-Main"]}"><a href="${item.ReportURL}" target="_blank"><div class="ms-md2  ${styles["bi-holder"]}" style='background-image:url("${icon}");'><div class='${styles["app-text"]}'>${item.Title}</div></div></a></div>
            `
         }
        
     });
     // Make Sure to Switch the nintex form to Prod when ready for Prod
        html += `
              </div>
            </div>
          </div>

          <div style="text-align: right">
            
            <div class="${styles.addButtonCont}"><a class="${styles.addButton}" href="https://winnebagoind-3c446e3a5fb2a1.sharepoint.com/sites/elt-team/FormsApp/NFLaunch.aspx?SPAppWebUrl=https://winnebagoind-3c446e3a5fb2a1.sharepoint.com/sites/elt-team/FormsApp&amp;SPHostUrl=https://winnebagoind.sharepoint.com/sites/elt-team&amp;remoteAppUrl=https://formso365.nintex.com&amp;ctype=0x0100077D5E1317F1E7428311EB8C605DA69E&amp;wtg=/NintexFormXml/0x0100077D5E1317F1E7428311EB8C605DA69E_2cd233d5-5c67-40a6-ad93-c6e0263ec769/&amp;mode=0&List=2cd233d5-5c67-40a6-ad93-c6e0263ec769&Source=https://winnebagoind.sharepoint.com/sites/elt-team/Lists/BIReports/AllItems.aspx&ContentTypeId=0x0100077D5E1317F1E7428311EB8C605DA69E&RootFolder=">
            Add an Item
            </a></div>
          
            <div class="${styles.addButtonCont}"><a class="${styles.addButton}" href="https://winnebagoind.sharepoint.com/sites/elt-team/SitePages/My-Reports-and-Dashboards.aspx">
            Remove an Item
            </a></div>
                        
          </div>
              
          `

     const listContainer: Element = this.domElement.querySelector('#spListContainer');
     listContainer.innerHTML = html;
     
   
   }
  public render(): void {
    this.domElement.innerHTML = `
      
      
      <div id="spListContainer"></div>
      `;

      this._renderListAsync();
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


