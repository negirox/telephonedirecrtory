import { MSGraphClientV3 } from "@microsoft/sp-http";
import { IUserProperties } from "./IUserProperties";
/**
 * Service to declare the methods
 */
export interface IMSGraphService{
    getUserProperties(email:string,context:MSGraphClientV3):Promise<IUserProperties[]>;
    getUserPropertiesBySearch(searchFor:string,client:MSGraphClientV3):Promise<IUserProperties[]>;
    getUserPropertiesByFirstName(searchFor:string,client:MSGraphClientV3):Promise<IUserProperties[]>;
    getUserPropertiesByLastName(searchFor:string,client:MSGraphClientV3):Promise<IUserProperties[]>;
}