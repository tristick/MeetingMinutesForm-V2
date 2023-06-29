import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IMeetingMinutesFormProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context:WebPartContext;
  siteUrl:string;
  custSiteUrl: string
  
}
export interface ICustomer {

  Title:string;

}
