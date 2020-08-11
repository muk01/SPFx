import {SPHttpClient} from '@microsoft/sp-http';

export interface IHelloWorldProps {
  ListName: string;
  spHttpClient: SPHttpClient;
  siteURL: string;
}
