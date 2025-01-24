declare interface IGreetingsWebPartStrings {
  PropertyPaneDescription: string;
  UserName: string;
  UserNameFieldLabel: string;
  UserNameFieldFirstName: string;
  UserNameFieldLastName: string;
  UserNameFieldFullName: string;
  UserNameFieldNoName: string;

  GreetingsFieldLabel:string;
  GreetingsTimeOfDayLabel:string;
  GreetingsWelcomeLabel:string;
  GreetingsHelloLabel:string;
  GreetingsHiLabel:string;
  GreetingsTimeOfDayMorning:string;
  GreetingsTimeOfDayAfternoon:string;
  GreetingsTimeOfDayEvening:string;
  GreetingsWelcomeMessage:string;
  GreetingsHelloMessage:string;
  GreetingsHiMessage:string;

  GreetingsCustomMessageLabel:string;
  GreetingsCustomMessageDescription:string;

  SupportMessage:string;

  FeedbackGroup:string;
}

declare module 'GreetingsWebPartStrings' {
  const strings: IGreetingsWebPartStrings;
  export = strings;
}
