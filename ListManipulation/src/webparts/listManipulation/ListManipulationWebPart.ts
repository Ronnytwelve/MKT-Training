import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
//import { escape } from '@microsoft/sp-lodash-subset';

//import styles from './ListManipulationWebPart.module.scss';
import * as strings from 'ListManipulationWebPartStrings';
import { ISPHttpClientOptions, SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface IListManipulationWebPartProps {
  description: string;
}

export default class ListManipulationWebPart extends BaseClientSideWebPart<IListManipulationWebPartProps> {

  //private _isDarkTheme: boolean = false;
  //private _environmentMessage: string = '';

  public render(): void {
    this.domElement.innerHTML = `
    <div>
    <div>
    <table border='5' bgcolor='aqua'>
    <tr>
    <td>Please Enter Software ID </td>
    <td><input type='text' id='txtID' />
    <td><input type='submit' id='btnRead' value='Read Details' />
    </td>
    </tr>
      <tr>
      <td>Software Title</td>
      <td><input type='text' id='txtSoftwareTitle' />
      </tr>
      <tr>
      <td>Software Name</td>
      <td><input type='text' id='txtSoftwareName' />
      </tr>
      <tr>
      <td>Software Vendor</td>
      <td>
      <select id="ddlSoftwareVendor">
        <option value="Microsoft">Microsoft</option>
        <option value="Sun">Sun</option>
        <option value="Oracle">Oracle</option>
        <option value="Google">Google</option>
      </select>  
      </td>
      </tr>
      <tr>
      <td>Software Version</td>
      <td><input type='text' id='txtSoftwareVersion' />
      </tr>
      <tr>
      <td>Software Description</td>
      <td><textarea rows='5' cols='40' id='txtSoftwareDescription'> </textarea>
      </td>
      </tr>
      <tr>
      <td colspan='2' align='center'>
      <input type='submit'  value='Insert Item' id='btnSubmit' />
      <input type='submit'  value='Update' id='btnUpdate' />
      <input type='submit'  value='Delete' id='btnDelete' />      
      </td>
    </table>
    </div><br/><br/><br/>
    <div id="divStatus"/>
    </div>`;
    this.bindEvents();
  }


  private bindEvents(): void {
    this.domElement.querySelector('#btnSubmit')!.addEventListener('click', () => { this.addListItem(); });
    this.domElement.querySelector('#btnRead')!.addEventListener('click', () => { this.readListItem(); });
    /* this.domElement.querySelector('#btnUpdate')!.addEventListener('click', () => { this.updateListItem(); });
    this.domElement.querySelector('#btnDelete')!.addEventListener('click', () => { this.deleteListItem(); }); */ 
  }

  //Fonction pour ajouter des éléments dans la liste MaListe qui a été créée
  private addListItem(): void {
    
    var softwaretitle = (document.getElementById("txtSoftwareTitle") as HTMLInputElement).value;
    var softwarename = (document.getElementById("txtSoftwareName") as HTMLInputElement).value;
    var softwareversion = (document.getElementById("txtSoftwareVersion") as HTMLInputElement).value;
    var softwarevendor = (document.getElementById("ddlSoftwareVendor") as HTMLInputElement).value;
    var softwareDescription = (document.getElementById("txtSoftwareDescription") as HTMLInputElement).value;

    //alert(softwareDescription);

    const siteurl: string = this.context.pageContext.site.absoluteUrl + "/_api/web/lists/getbytitle('MaListe')/items";

    const itemBody: any = {
      "Title": softwaretitle,
      "SoftwareVendor": softwarevendor,
      "SoftwareDescription": softwareDescription,
      "SoftwareName": softwarename,
      "SoftwareVersion": softwareversion,
     
    };

    const spHttpClientOptions: ISPHttpClientOptions = {
      "body": JSON.stringify(itemBody)
    };

    this.context.spHttpClient.post(siteurl, SPHttpClient.configurations.v1, spHttpClientOptions).then((response: SPHttpClientResponse) => {
     
      if (response.status === 201) {
        let statusmessage: Element = this.domElement.querySelector('#divStatus')!;
        statusmessage.innerHTML = "L'élément de liste a été créé avec succès.";
        this.clear();

      }
      else {
        let statusmessage: Element = this.domElement.querySelector('#divStatus')!;
        statusmessage.innerHTML = "Une erreur s'est produite " + response.status + " - " + response.statusText;
      }
      
    });
    
  }

  private readListItem(): void {

  }

  //Fonction pour vider les champs du formulaire
  private clear(): void {
    (document.getElementById("txtSoftwareTitle") as HTMLInputElement).value= '';
    (document.getElementById("ddlSoftwareVendor") as HTMLInputElement).value= 'Microsoft';
    (document.getElementById("txtSoftwareDescription") as HTMLInputElement).value= '';
    (document.getElementById("txtSoftwareVersion") as HTMLInputElement).value= '';
    (document.getElementById("txtSoftwareName") as HTMLInputElement).value= '';

  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      //this._environmentMessage = message;
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    //this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
