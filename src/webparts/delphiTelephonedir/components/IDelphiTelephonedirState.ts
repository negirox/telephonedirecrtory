import { IColumn } from "office-ui-fabric-react/lib/DetailsList";

export interface IDelphiTelephonedirState {
    loading: boolean;
    columns: IColumn[];
    selectedKey: string;
}
