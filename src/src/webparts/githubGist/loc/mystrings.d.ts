declare interface IGithubGistStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  TitleFieldLabel: string;
  GistIdFieldLabel: string;
  GistFileFieldLabel: string;
  WarningGistPropertyNotSet: string;
}

declare module 'githubGistStrings' {
  const strings: IGithubGistStrings;
  export = strings;
}
