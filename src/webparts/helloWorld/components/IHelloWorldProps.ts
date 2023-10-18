/* eslint-disable @typescript-eslint/no-explicit-any */
export interface IHelloWorldProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: any;
  showConsoleLogs:()=>void;
}
