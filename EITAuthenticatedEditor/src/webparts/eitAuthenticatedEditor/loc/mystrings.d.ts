declare interface IEitAuthenticatedEditorWebPartStrings {
  WebPartTitle: string;
  WebPartTitleDescription: string;
  PropertyPaneDescription: string;
  DescriptionFieldLabel: string;

  // Web Part Properties
  TemplateUrl: string;
  TemplateUrlDescription: string;
  Audience: string;
  AudienceDescription: string;

  NoTemplateProvided: string;
  TemplateError: string;

  LoadingText: string;
  SavingText: string;
  SuccessText: string;
  ErrorText: string;
}

declare module 'EitAuthenticatedEditorWebPartStrings' {
  const strings: IEitAuthenticatedEditorWebPartStrings;
  export = strings;
}
