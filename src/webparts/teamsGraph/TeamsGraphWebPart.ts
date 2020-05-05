import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import { MSGraphClient } from '@microsoft/sp-http';
// import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
declare const Buffer
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

 html+= `<br/><textarea id="msg" name="msg" rows="6" cols="80" placeholder="Enter message" class="${styles.textareas}"></textarea>`;
 html+= '<br/><input type="file" name="file" id="myfileinput"  multiple>';
 html+=`<ul class="" style='list-style-type:none;' ><li class="${styles.listItem}" >`;
 html+= '<br/><span class="ms-font-1"><input type="checkbox" name="all" id="all">Select All Teams </input></span>';
 html+= `</li></ul>`;
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
        html+= `<center><input type="button" id="btn" value="Send" class="${styles.button}"> </center><br/>`;
        let htmlElement: Element = this.domElement.querySelector('#msteamcontainer');
        htmlElement.innerHTML = html;

        //console.log(myhash);

        this._setButtonEventHandlers();
}

private _setButtonEventHandlers():void{
  //const webPart: WpConfigureApplicationCustomizerWebPart = this;

  this.domElement.querySelector('#btn').addEventListener('click',()=>{
      this._formsubmission().then(ress=>{
        //alert("Response: "+ress.message);
        let htmlElement: Element = this.domElement.querySelector('#status');
        htmlElement.innerHTML = "<span style='color:red;'>Message successfully sent</span>";
      });
  })

  this.domElement.querySelector('#all').addEventListener('click',()=>{
    var element = <HTMLInputElement> this.domElement.querySelector('#all');
    var isChecked = element.checked?true:false;
    var checkboxes : NodeListOf<Element> = this.domElement.querySelectorAll('input[type="checkbox"]');
    for (var checkbox of checkboxes as any) {
      checkbox.checked = isChecked;
  }
  });
}

private async _formsubmission():Promise<any>{
  let htmlElement: Element = this.domElement.querySelector('#status');
  htmlElement.innerHTML = "<span style='color:red;'>Please wait while message is sending</span>";
  var message:HTMLTextAreaElement = this.domElement.querySelector('#msg');
  var checkboxes: NodeListOf<Element> = this.domElement.querySelectorAll('input[type="checkbox"]:checked');

  console.log("Checkbox: ",checkboxes.item(0));

  var fileInput:HTMLInputElement = this.domElement.querySelector("#myfileinput") ;

  // files is a FileList object (similar to NodeList)
  var files = fileInput.files;

  // object for allowed media types
  var accept = {
    binary : ["image/png", "image/jpeg"],
    text   : ["text/plain", "text/css", "application/xml", "text/html"]
  };

  var file;
  let channelId: string[] = [];
  var promise2 = null;
  var fileContext = "";
 var fnames = []; 
   for ( var checkbox of checkboxes as any) {
     if(checkbox.name!=='all'){
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
              
                if(files.length>0){

              //upload files
              var siteApi = "/sites/root:/sites/"+checkbox.name;

              var clientSPContext:string = "";

              this.context.msGraphClientFactory
              .getClient()
              .then((client:MSGraphClient):void=>{
                  client
                  .api(siteApi)
                  .version("beta")
                  .get((error, messages: any, rawResponse?: any) => {

                    var siteId:string = messages.id;

                    clientSPContext = messages.webUrl;

                    console.log("Site Id: ",siteId);
                    var driveApi = "/sites/"+siteId+"/drives";
                    this.context.msGraphClientFactory
                    .getClient()
                    .then((client:MSGraphClient):void=>{
                        client
                        .api(driveApi)
                        .version("beta")
                        .get((error, messages: any, rawResponse?: any) => {
                            var driveId:string = messages.value[0].id;
                            console.log("Drive Id: ",driveId);
                            var fileUrl ="";
                            var p1,p2,p3,p4;
                            fnames=[];
                            
                            p1 = new Promise(resolvesOFor=>{ 
                              for (var i = 0; i < files.length; i++) {
                              p2 = new Promise(resolvesIFor=>{
                                file = files[i];
                                fnames.push(file.name);
                                var uploadApi = driveApi + "/"+driveId+"/root:/General/"+file.name+":/content";

                                p3=this._getFileRead(file).then(response=>{
                              
                                console.log("Result type: ",typeof( response));
                                //console.log("Result msgs type: ",typeof( Messagess));
                                console.log("msgs: ", response);
                                //uploadMessage = uploadMessage.substring(1, uploadMessage.length() - 1)
                                
                                p4=this.context.msGraphClientFactory
                                .getClient()
                                .then((client:MSGraphClient):void=>{
                                    client
                                    .api(uploadApi)
                                    .version("beta")
                                    .put(response)
                                    .then((res)=>{
                                      fileUrl = fileUrl+" "+res.webUrl;
                                      console.log("Weburl: ",fileUrl);
                                      
                                    })
                                    
                                  })
                                })
                              resolvesIFor();                            
                          })
                          resolvesOFor()
                        }
                        //resolvesssss(fileUrl)
                      })
                      Promise.all([p1,p2,p3,p4]).then((res)=>{
                        console.log("urls: ",res)
                        console.log("urlss: ",fileUrl)
                        fileContext="";

                        var chatMessage = {
                          subject: null,
                          body: {
                              contentType: "html",
                              content: message.value + fnames.map((ele)=>{
                                var temp=clientSPContext +"/"+"Shared%20Documents/General/"+ele;
                                return "<br/><a href='"+temp+"'>"+ele+"</a>";
                              })
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
                  })
                  }); 
              }
              else{
              
              var chatMessage = {
                subject: null,
                body: {
                    contentType: "html",
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
              }
            })            
          })    
        })
    }
  }
}

  private  _getFileRead(file):Promise<any>{ 
     return new Promise<any>((resolve,reject) =>{
      var fr = new FileReader();  
      fr.onload = () => {    
        console.log("data: ",fr.result)
        resolve(fr.result )
      };
      fr.readAsArrayBuffer(file);
    })
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
