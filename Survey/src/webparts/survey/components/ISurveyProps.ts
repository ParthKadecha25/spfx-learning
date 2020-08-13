import { SPHttpClient } from "@microsoft/sp-http";

export interface ISurveyProps {
  description: string;
  surveyTitle : string;
  surveyDescription? : string;
  listNameSurvey: string;
  spHttpClient: SPHttpClient;
  siteUrl: string;
  adminMode : boolean,
  listNameSurveyResponse : string;
}
