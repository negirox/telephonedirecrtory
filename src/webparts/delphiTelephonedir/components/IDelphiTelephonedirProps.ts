import { MSGraphService } from "../../../Services/MSGraphService";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { DisplayMode } from '@microsoft/sp-core-library';
export interface IDelphiTelephonedirProps {
  description: string;
  MSGraphServiceInstance: MSGraphService;
  context: WebPartContext;
  MsGraphClient: MSGraphClientV3;
  DisplayMode: DisplayMode;
  WebpartTitle: string;
  updateProperty: (value: string) => void;
  isDisplayName:boolean,
  isEmail:boolean,
  ismobilePhone:boolean,
  isJobTitle:boolean,
  isOfficeLocation:boolean,
  isbusinessPhone:boolean,
  showBorder:boolean,
}
