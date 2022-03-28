declare interface INotifyToTeamsGroupCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'NotifyToTeamsGroupCommandSetStrings' {
  const strings: INotifyToTeamsGroupCommandSetStrings;
  export = strings;
}
