import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './VisitorsWebPartWebPart.module.scss';
import * as strings from 'VisitorsWebPartWebPartStrings';
import {  
  SPHttpClient, SPHttpClientResponse  
} from '@microsoft/sp-http';
import {  
  Environment,  
  EnvironmentType  
} from '@microsoft/sp-core-library'; 
import {SPComponentLoader} from '@microsoft/sp-loader';
import * as jquery from "jquery";
require("jqueryui");
require('datatables');


export interface IVisitorsWebPartWebPartProps {
  description: string;
}
export interface ISPLists {  
  value: ISPList[];
}

//Interface used to set properties for visitor item
export interface ISPList {  
  Title: string;  
  PhoneNumber: string;  
  VisitReason: string;  
  VisitorStatus: string;  
  DateTime:string;
  OutDateTime:string;
}

export default class VisitorsWebPartWebPart extends BaseClientSideWebPart<IVisitorsWebPartWebPartProps> {
  public existingVisitorItems;
  public constructor() {
    super();
    //Load jquery-ui.min.css: required for jquery UI datepicker
    SPComponentLoader.loadCss("//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.min.css");
    //Load jquery.dataTables.min.css: required for jquery datatables
    SPComponentLoader.loadCss("//cdn.datatables.net/1.10.18/css/jquery.dataTables.min.css");
  }
  public render(): void {
    this.domElement.innerHTML = `
    <div id="spListContainer" /> `;

    //Load existing visitor items in datatable
    this.getExistingVisitorData();
  }

  //Get existing training data from SharePoint list
  private getListData(): Promise<ISPLists> {  
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists/GetByTitle('VisitorsInformation')/Items?$select=ID,Title,PhoneNumber,VisitReason,VisitorStatus,DateTime,OutDateTime`, SPHttpClient.configurations.v1)  
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          return response.json();
        } else {
          alert("An error occured while fetching existing trainings. Please contact your administrator!");
          console.log(response.statusText);
        }
      });  
  }

  //Get existing visitor data and bind to datatable
  private getExistingVisitorData(): void {  
       this.getListData()  
        .then((response) => {
          if(response) {
            var finalItems: ISPList[] = [];
            response.value.forEach((item: any) =>{
             
              var listItem:ISPList = {      
                Title: item.Title, 
                PhoneNumber: item.PhoneNumber,  
                VisitReason: item.VisitReason,  
                VisitorStatus: item.VisitorStatus, 
                DateTime: item.DateTime ? item.DateTime : "",
                OutDateTime: item.OutDateTime ? item.OutDateTime : ""
              };
              finalItems.push(listItem);
              this.existingVisitorItems = finalItems;
            });
            this.bindVisitorsToDatatable(finalItems);
          }
        });  
  }

  //Bind existing visitor data to datatable
  private bindVisitorsToDatatable(items: ISPList[]): void {  
    let html: string = "";
    if(items.length) {
      html += `<table id="tbVisitors" class="${styles.Vtable}">`;  
      html += `<thead><tr><th>Visitor Name</th><th>Phone Number</th></th><th>Reason For Visit</th><th>Visitor Status</th><th>Checked In Time</th><th>Checked Out Time</th></tr></thead><tbody>`;
      items.forEach((item: ISPList) => {  
        html += `  
            <tr>  
            <td>${item.Title}</td>  
            <td>${item.PhoneNumber}</td> 
            <td>${item.VisitReason}</td> 
            <td>${item.VisitorStatus}</td>  
            <td>${item.DateTime}</td>
            <td>${item.OutDateTime}</td>  
            </tr>  
            `;  
      });  
      html += `</tbody></table>`;
    } else {
      html += "<p>No existing visitors found.";
    }
    const listContainer: Element = this.domElement.querySelector('#spListContainer');  
    listContainer.innerHTML = html;

    //Bind to datatable
    var table = jquery('#tbVisitors').DataTable({
      "orderCellsTop": true,
      "fixedHeader": true,
      "pageLength": 5
    });

    // Setup - add a text input to each footer cell
    jquery('#tbVisitors thead tr').clone(true).appendTo( '#tbVisitors thead' );
    jquery('#tbVisitors thead tr:eq(1) th').each( function (i) {
        var title = jquery(this).text();
        jquery(this).html( '<input type="text" placeholder="Search '+title+'" />' );
 
        jquery( 'input', this ).on( 'keyup change', function () {
            if ( table.column(i).search() !== this.value ) {
                table
                    .column(i)
                    .search( this.value )
                    .draw();
            }
        } );
    } );
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
