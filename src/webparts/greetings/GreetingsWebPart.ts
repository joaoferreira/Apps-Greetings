import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import * as strings from 'GreetingsWebPartStrings';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneLink,
  PropertyPaneChoiceGroup
} from '@microsoft/sp-property-pane';

import styles from './GreetingsWebPart.module.scss';

export interface IGreetingsWebPartProps {
  userName: string;
  userNameFormat:string;
  greetingsMessage:string;
  greetingsCustomMessage:string;
}

export default class GreetingsWebPart extends BaseClientSideWebPart<IGreetingsWebPartProps> {
  public render(): void {    
    const userName = this.getUserNameFormat(this.properties.userName);
    const message = this.getGreetingsMessage(this.properties.greetingsMessage)
    const greeting = `${message} ${userName}`
    this.domElement.innerHTML = `<h2 class="${ styles.greetings }">${greeting}</h2>`;
  }

  protected getUserNameFormat(userName:string): string{
    let name:string; 
    switch(this.properties.userNameFormat) { 
      case 'fullname': { 
         name = userName;
         break; 
      } 
      case 'firstname': { 
         name = userName.split(' ')[0];
         break; 
      } 
      case 'lastname': { 
        name = userName.split(' ')[userName.split(' ').length-1];
        break; 
      } 
      case 'noname': { 
        name = "";
        break; 
      } 
      default: { 
        name = userName;
        break; 
      } 
   } 
    return name;
  }

  protected getGreetingsMessage(format:string): string{
    let greeting:string; 
    switch(this.properties.greetingsMessage) { 
      case 'timeofday': { 
        // Get the current date and time
        const currentDate = new Date();
        // Get the current hour from the current date
        const currentHour = currentDate.getHours();
        // Check the current hour and choose a greeting accordingly
        if (currentHour < 12) {
          greeting = strings.GreetingsTimeOfDayMorning;
        } else if (currentHour < 18) {
          greeting = strings.GreetingsTimeOfDayAfternoon;
        } else {
          greeting = strings.GreetingsTimeOfDayEvening;
        }
        break; 
      } 
      case 'welcome': { 
        greeting = strings.GreetingsWelcomeMessage;
        break; 
      } 
      case 'hello': { 
        greeting = strings.GreetingsHelloMessage;
        break; 
      }
      case 'hi': { 
        greeting = strings.GreetingsHiMessage;
        break; 
      }  
      default: { 
        greeting = strings.GreetingsWelcomeMessage;
        break; 
      } 
    }
    
    //Replace the coma if the user decided to not display the user name
    if(this.properties.userNameFormat === "noname"){
      greeting = greeting.replace(',','');
    }

    //Get the custom message if defined in the text field
    if(this.properties.greetingsCustomMessage !== "") {
      greeting = this.properties.greetingsCustomMessage; 
    }

    return greeting;
  }

  protected async onInit(): Promise<void> {
    //Init PnP JS and get the current user name
    const sp = spfi().using(SPFx(this.context));
    const user = await sp.web.currentUser();
    this.properties.userName = user.Title;  
    return super.onInit();
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
              groupFields: [
                PropertyPaneChoiceGroup('userNameFormat',{
                  label: strings.UserNameFieldLabel,
                  options: [
                    { key: 'fullname', text: strings.UserNameFieldFullName },
                    { key: 'firstname', text: strings.UserNameFieldFirstName },
                    { key: 'lastname', text: strings.UserNameFieldLastName },
                    { key: 'noname', text: strings.UserNameFieldNoName }
                  ]
                }),
                PropertyPaneChoiceGroup('greetingsMessage',{
                  label: strings.GreetingsFieldLabel,
                  options: [
                    { key: 'timeofday', text: strings.GreetingsTimeOfDayLabel },
                    { key: 'welcome', text: strings.GreetingsWelcomeLabel },
                    { key: 'hello', text: strings.GreetingsHelloLabel },
                    { key: 'hi', text: strings.GreetingsHiLabel }
                  ]
                }),
                PropertyPaneTextField('greetingsCustomMessage', {
                  label: strings.GreetingsCustomMessageLabel,
                  description: strings.GreetingsCustomMessageDescription
                })
              ]              
            },
            {
              groupName: strings.FeedbackGroup,
              groupFields: [
                PropertyPaneLink('supportLink', {
                href: 'http://handsontek.net/apps',
                text: strings.SupportMessage,
                target: '_blank'})
              ]
            }
          ]
        }
      ]
    };
  }
}
