import { WebPartContext } from "@microsoft/sp-webpart-base";
import { MSGraphService } from "../../../../Services/MSGraphService";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { IColumn } from "office-ui-fabric-react/lib/DetailsList";
export interface ByEmailProps {
    context: WebPartContext;
    MSGraphServiceInstance: MSGraphService;
    MSGraphClient: MSGraphClientV3;
    columns: IColumn[];
}