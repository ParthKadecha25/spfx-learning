import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-webpart-base';

import * as strings from 'SurveyWebPartStrings';
import Survey from './components/Survey';
import { ISurveyProps } from './components/ISurveyProps';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface ISurveyWebPartProps {
  surveyAdminGroupName: string;
}

export default class SurveyWebPart extends BaseClientSideWebPart<ISurveyWebPartProps> {

  private SurveyAdminGroupDDLOptions: IPropertyPaneDropdownOption[] =[];  

  public render(): void {
    if(this.properties.surveyAdminGroupName){
      let page = this;
      this.checkUserIsMemberOfGroup(this.properties.surveyAdminGroupName).then(function(isMember : boolean){
        const element: React.ReactElement<ISurveyProps > = React.createElement(
          Survey,
          {
            spHttpClient: page.context.spHttpClient,
            siteUrl: page.context.pageContext.web.absoluteUrl,
            listName: "Survey",
            adminMode : isMember
          }
        );
        ReactDom.render(element, page.domElement);
      });
    }
    else{
      this.domElement.innerHTML = `<h1>Please configure the webpart properties!</h1>`;
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

   
  protected onPropertyPaneConfigurationStart(): Promise<void> {  
    let page = this;
    if(this.SurveyAdminGroupDDLOptions.length > 0) return;
    return this.getSiteGrups().then(function(){      
      page.context.propertyPane.refresh();
    });
 } 

 protected getSiteGrups(): Promise<void> {
  return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/sitegroups`, SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse) => {
      return response.json();
    }).then((response) => {
      let items = response.value;      
      items.forEach((item) => {
        this.SurveyAdminGroupDDLOptions.push({key: item["Id"], text: item["Title"]});
      });
    });
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration  {
    return {
      pages: [
        {
          header: {
            description: "Survey Configuration"
          },
          groups: [
            {
              groupName: "Survey Administration",
              groupFields: [
                PropertyPaneDropdown('surveyAdminGroupName', {
                  label: "Survey Admin Group:",
                  options : this.SurveyAdminGroupDDLOptions,
                  disabled : false
                })
              ]
            }
          ]
        }
      ]
    };
  }

  protected checkUserIsMemberOfGroup(groupName : string) : Promise<boolean>{
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/sitegroups/getById('` + groupName + `')/Users?$filter=Email eq '` + this.context.pageContext.user.email + `'`, SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse) => {
      return response.json();
    }).then((response) => {
      let items = response.value;
      return (items.length > 0)
    });
  }
}
