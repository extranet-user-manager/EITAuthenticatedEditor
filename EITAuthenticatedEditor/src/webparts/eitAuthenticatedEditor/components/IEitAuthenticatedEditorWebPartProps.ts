import { HttpClient, AadTokenProvider } from "@microsoft/sp-http";

export interface IEitAuthenticatedEditorWebPartProps {
    WebPartTitle?: string;
    SiteCollectionTitle: string;
    SiteCollectionURL: string;
    HttpClient: HttpClient;
    AadTokenProvider: AadTokenProvider;
   
    spPageContextInfo: boolean;
    teamsContext: boolean;
    script: string;
    disposeScript: string;
  
    TemplateUrl: string;
    TemplateHtml: string;
    Audience: string;
    Resources: string;
  }