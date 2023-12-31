import { IMSGraphService } from "./IMSGraphService";
import { IUserProperties } from "./IUserProperties";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { Log } from "@microsoft/sp-core-library";
const LOG_SOURCE = "MSGraphService";
export class MSGraphService implements IMSGraphService{

    public async getUserProperties(email:string,client:MSGraphClientV3):Promise<IUserProperties[]>{
        const userProperties:IUserProperties[] = [];
        try {
            //let client:MSGraphClient = await context.msGraphClientFactory.getClient().then();
            const endPoint = `/Users/${email}`;
            const response = await client.api(`${endPoint}`).version("v1.0").get();
            if(response){
                userProperties.push({
                   businessPhone:response.businessPhones[0],
                   displayName:response.displayName,
                   email:response.mail,
                   JobTitle:response.jobTitle,
                   OfficeLocation:response.officeLocation,
                   mobilePhone:response.mobilePhone,
                   preferredLanguage:response.preferredLanguage
               });
           }
        } catch (error) {
            console.log(error);
            Log.error(LOG_SOURCE+"getUserProperties():",error);
        }
        return userProperties;
    }

    public async getUserPropertiesByLastName(searchFor:string,client:MSGraphClientV3):Promise<IUserProperties[]>{
        const userProperties:IUserProperties[] = [];
        try {
            const res = await client.api("users")
            .version("v1.0")
            .filter(`(startswith(surname,'${escape(searchFor)}'))`).get();
            if(res.value.length !== 0){
                // eslint-disable-next-line @typescript-eslint/no-unused-vars, @typescript-eslint/no-explicit-any
                res.value.map((_userProperty:any,_index:any)=>{
                    if(_userProperty.mail !== null){
                        userProperties.push({
                            businessPhone:_userProperty.businessPhones[0],
                            displayName:_userProperty.displayName,
                            email:_userProperty.mail,
                            JobTitle:_userProperty.jobTitle,
                            OfficeLocation:_userProperty.officeLocation,
                            mobilePhone:_userProperty.mobilePhone,
                            preferredLanguage:_userProperty.preferredLanguage
                        });
                    }
                });
            }
        } catch (error) {
            console.log(error);
            Log.error(LOG_SOURCE+"getUserPropertiesByLastName():",error);
        }
        return userProperties;

    }

    public async getUserPropertiesByFirstName(searchFor:string,client:MSGraphClientV3):Promise<IUserProperties[]>{
        const userProperties:IUserProperties[] = [];
        try {
            const res = await client.api("users")
            .version("v1.0")
            .filter(`(startswith(givenName,'${escape(searchFor)}'))`).get();
            if(res.value.length !== 0){
                // eslint-disable-next-line @typescript-eslint/no-unused-vars, @typescript-eslint/no-explicit-any
                res.value.map((_userProperty:any,_index:any)=>{
                    if(_userProperty.mail !== null){
                        userProperties.push({
                            businessPhone:_userProperty.businessPhones[0],
                            displayName:_userProperty.displayName,
                            email:_userProperty.mail,
                            JobTitle:_userProperty.jobTitle,
                            OfficeLocation:_userProperty.officeLocation,
                            mobilePhone:_userProperty.mobilePhone,
                            preferredLanguage:_userProperty.preferredLanguage
                        });
                    }
                });
            }
        } catch (error) {
            console.log(error);
            Log.error(LOG_SOURCE+"getUserPropertiesBySearch():",error);
        }
        return userProperties;

    }
    public async getUserPropertiesBySearch(searchFor:string,client:MSGraphClientV3):Promise<IUserProperties[]>{
        const userProperties:IUserProperties[] = [];
        try {
            const res = await client.api("users")
            .version("v1.0")
            .filter(`(startswith(displayName,'${escape(searchFor)}'))`).get();
            if(res.value.length !== 0){
                // eslint-disable-next-line @typescript-eslint/no-unused-vars, @typescript-eslint/no-explicit-any
                res.value.map((_userProperty:any,_index:any)=>{
                    userProperties.push({
                        businessPhone:_userProperty.businessPhones[0],
                        displayName:_userProperty.displayName,
                        email:_userProperty.mail,
                        JobTitle:_userProperty.jobTitle,
                        OfficeLocation:_userProperty.officeLocation,
                        mobilePhone:_userProperty.mobilePhone,
                        preferredLanguage:_userProperty.preferredLanguage
                    });
                });
            }
        } catch (error) {
            console.log(error);
            Log.error(LOG_SOURCE+"getUserPropertiesBySearch():",error);
        }
        return userProperties;

    }
}
