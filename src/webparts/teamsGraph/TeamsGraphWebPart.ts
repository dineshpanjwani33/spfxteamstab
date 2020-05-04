import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import { MSGraphClient } from '@microsoft/sp-http';
// import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

import styles from './TeamsGraphWebPart.module.scss';
import * as strings from 'TeamsGraphWebPartStrings';

export interface ITeamsGraphWebPartProps {
  description: string;
}

export interface IHash {
    [details: string] : string;
} 

let myhash: IHash = {};



export default class TeamsGraphWebPart extends BaseClientSideWebPart <ITeamsGraphWebPartProps> {



private _renderList(messages:any):void{

 let html:string = '';

 html+= '<br/><textarea id="msg" name="msg" rows="7" cols="80"></textarea>';
  messages.value.map((item: any)=>{
          html += `
          <ul class="${styles.list}" >
            <li class="${styles.listItem}" >
              <span class="ms-font-1"><input type='checkbox' name="${item.displayName}" id="${item.displayName}">${item.displayName} </input></span>
            </li>
          </ul>
          `;

          myhash[item.displayName] = item.id;          

        });
        html+= '<center><input type="button" id="btn" value="Send"> </center><br/>';
        let htmlElement: Element = this.domElement.querySelector('#msteamcontainer');
        htmlElement.innerHTML = html;

        //console.log(myhash);

        this._setButtonEventHandlers();
}

private _setButtonEventHandlers():void{
  //const webPart: WpConfigureApplicationCustomizerWebPart = this;

  this.domElement.querySelector('#btn').addEventListener('click',()=>{
      this._formsubmission().then((response)=>{
        //alert("Message successfully sent...");
        let htmlElement: Element = this.domElement.querySelector('#status');
        htmlElement.innerHTML = "<span style='color:red;'>Message successfully sent</span>";
      });
  })
}

private async _formsubmission():Promise<any>{
  let htmlElement: Element = this.domElement.querySelector('#status');
  htmlElement.innerHTML = "<span style='color:red;'>Please wait while message is sending</span>";
  var message:HTMLTextAreaElement = this.domElement.querySelector('#msg');
  var checkboxes: NodeListOf<Element> = this.domElement.querySelectorAll('input[type="checkbox"]:checked');
  let channelId: string[] = [];
   for ( var checkbox of checkboxes as any) {
    console.log("Team Name: "+checkbox.name+" : "+myhash[checkbox.name]);
    //var teamId = myhash[checkbox.name];
    var apiCall = "/teams/"+myhash[checkbox.name]+"/channels";
    await new Promise(resolve => {
      this.context.msGraphClientFactory
      .getClient()
      .then((client:MSGraphClient):void=>{
          client
          .api(apiCall)
          .version("beta")
          .get((error, messages: any, rawResponse?: any) => {
            console.log("Team Id: "+myhash[checkbox.name]);
            console.log(" Channel Id: "+messages.value[0].id + "Message: "+message.value);
            //channelId.push(messages.value[0].id);
            var msgApi = apiCall+"/"+messages.value[0].id+"/messages";
            var chatMessage = {
              subject: null,
              body: {
                  contentType: "text",
                  content: message.value
              }
            }
            this.context.msGraphClientFactory
              .getClient()
              .then((clients:MSGraphClient):void=>{
                  clients
                  .api(msgApi)
                  .version("beta")
                  .post(chatMessage)
              })
            resolve();
        })
            
    })
    
    })
  }

  //this._sendMessageToChannel(channelId,Message.value)

}

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.teamsGraph }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Welcome to Teams!</span>
              <p class="${ styles.description }">${escape(this.properties.description)}</p>
              </a>
            </div>
          </div>
          <div id="msteamcontainer"></div>
          <div id="status"></div>
        </div>
      </div>`;
        this._renderMsgForm();
  }

  private _renderMsgForm():void{
    this.context.msGraphClientFactory
    .getClient()
    .then((client: MSGraphClient):void=>{
      client
      .api("/me/joinedTeams")
      .version("beta")
      .get((error, messages: any, rawResponse?: any) => {
       
        
        if(error){
          console.log("Error: "+error);
        }
          this._renderList(messages);
        
      });
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
