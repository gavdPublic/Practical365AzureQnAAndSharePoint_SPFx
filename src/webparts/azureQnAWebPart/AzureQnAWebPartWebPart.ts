import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

// This are the JS Libraries to make HTTP calls
import { 
  HttpClient, 
  HttpClientResponse, 
  IHttpClientOptions, 
} from '@microsoft/sp-http';

import styles from './AzureQnAWebPartWebPart.module.scss';
import * as strings from 'AzureQnAWebPartWebPartStrings';

export interface IAzureQnAWebPartWebPartProps {
  description: string;
}

export default class AzureQnAWebPartWebPart extends BaseClientSideWebPart<IAzureQnAWebPartWebPartProps> {

  protected hostUrl:     string = "https://prac365qnamakermaserati.azurewebsites.net";
  protected indexId:     string = "1fb7e07c-2752-4ef1-b768-55fc3a0c3cad";
  protected endpointKey: string = "beb53282-ca29-49c8-8db1-e32fb934e499";

  protected QnAUrl: string = this.hostUrl + "/qnamaker/knowledgebases/" + this.indexId + "/generateAnswer";

  protected getAnswers(): void {

    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    requestHeaders.append("Cache-Control", "no-cache");
    requestHeaders.append("Authorization", "EndpointKey " + this.endpointKey);

    // Gather the information from the form fields
    var strQuestion: string = (<HTMLInputElement>document.getElementById("txtQuestion")).value;

    // This are the options for the HTTP call
    const callOptions: IHttpClientOptions = {
      headers: requestHeaders,
      body: `{'question': '${strQuestion}', 'top': 1}`
    };

    // Create the responce object
    let responceAnswer: HTMLElement = document.getElementById("responseAnswer");
    let responceScore:  HTMLElement = document.getElementById("responseScore");
    let responceSource: HTMLElement = document.getElementById("responseSource");

    // And make a POST request to the Function
    this.context.httpClient.post(this.QnAUrl, HttpClient.configurations.v1, callOptions).then((response: HttpClientResponse) => {
       response.json().then((responseJSON: JSON) => {
          var myResponseText = JSON.stringify(responseJSON);
          var myResponseJson = JSON.parse(myResponseText);

          responceAnswer.innerText = myResponseJson.answers[0].answer;
          responceScore.innerText  = myResponseJson.answers[0].score;
          responceSource.innerText = myResponseJson.answers[0].source;
        })
        .catch ((response: any) => {
          let errorMessage: string = `Error calling ${this.QnAUrl} = ${response.message}`;
          responceAnswer.innerText = errorMessage;
        });
    });
  }

  // This is the interface
  public render(): void {
    this.domElement.innerHTML = `
    <div class="${styles.azureQnAWebPart}">
    <div class="${styles.container}">
      <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
        <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
          <span class="ms-font-xl ms-fontColor-white">Maserati QnA</span>
          <div class="${styles.controlRow}">
            <span class="ms-font-l ms-fontColor-white ${styles.controlLabel}">Question:</span>
            <textarea id="txtQuestion" rows="3" cols="50">Can a Maserati maintain its value?</textarea>
          </div>
          <div class="${styles.buttonRow}"></div>
          <button id="btnGetAnswers" class="${styles.button}">Get Answer</button>
          <div><span class="ms-font-l ms-fontColor-white ${styles.controlLabel}">Answer:</span>
          <div id="responseAnswer" class="${styles.resultRow}"></div></div>
          <div><span class="ms-font-l ms-fontColor-white ${styles.controlLabel}">Score:</span>
          <div id="responseScore" class="${styles.resultRow}"></div></div>
          <div><span class="ms-font-l ms-fontColor-white ${styles.controlLabel}">Source:</span>
          <div id="responseSource" class="${styles.resultRow}"></div></div>
        </div>
      </div>
    </div>
    </div>`;

    // The Event Handler for the Button  
    document.getElementById("btnGetAnswers").onclick = this.getAnswers.bind(this);
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
