import * as React from 'react';
import styles from './DelphiTelephonedir.module.scss';
import * as strings from 'DelphiTelephonedirWebPartStrings';
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import { ByFirstName } from "./ByFirstName/ByFirstName";
import { ByLastName } from "./ByLastName/ByLastName";
import { ByEmail } from "./ByEmail/ByEmail";
import { Pivot, PivotItem, PivotLinkFormat, PivotLinkSize } from 'office-ui-fabric-react/lib/Pivot';
import { IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { IDelphiTelephonedirProps } from './IDelphiTelephonedirProps';
import { IDelphiTelephonedirState } from './IDelphiTelephonedirState';
const columns: IColumn[] = [
  {
    key: 'column1',
    name: strings.DisplayName,
    isRowHeader: true,
    isSorted: true,
    isSortedDescending: false,
    sortAscendingAriaLabel: 'Sorted A to Z',
    sortDescendingAriaLabel: 'Sorted Z to A',
    fieldName: 'displayName',
    minWidth: 100,
    maxWidth: 300,
    isResizable: false
  },
  {
    key: 'column2',
    name: strings.Email,
    fieldName: 'email',
    isSorted: true,
    isSortedDescending: false,
    minWidth: 200,
    maxWidth: 300,
    isResizable: false
  },
  {
    key: 'column3',
    name: strings.MobilePhone,
    fieldName: 'mobilePhone',
    isSorted: true,
    isSortedDescending: false,
    minWidth: 200,
    maxWidth: 300,
    isResizable: false
  },
  {
    key: 'column5',
    name: strings.JobTitle,
    fieldName: 'JobTitle',
    isSorted: true,
    isSortedDescending: false,
    minWidth: 200,
    maxWidth: 300,
    isResizable: false
  },
  {
    key: 'column6',
    name: strings.OfficeLocation,
    fieldName: 'OfficeLocation',
    isSorted: true,
    isSortedDescending: false,
    minWidth: 100,
    maxWidth: 300,
    isResizable: true
  },
  {
    key: 'column7',
    name: strings.businessPhone,
    fieldName: 'businessPhone',
    isSorted: true,
    isSortedDescending: false,
    minWidth: 100,
    maxWidth: 300,
    isResizable: true
  }
];
export default class DelphiTelephonedir extends React.Component<IDelphiTelephonedirProps, IDelphiTelephonedirState> {
  constructor(props: IDelphiTelephonedirProps) {
    super(props);
    const columnToShow : IColumn[] =[];
    if(this.props.isDisplayName){
      this.addToColumns(columnToShow,'displayName');
    }
    if(this.props.isEmail){
      this.addToColumns(columnToShow,'email');
    }
    if(this.props.ismobilePhone){
      this.addToColumns(columnToShow,'mobilePhone');
    }
    if(this.props.isJobTitle){
      this.addToColumns(columnToShow,'JobTitle');
    }
    if(this.props.isOfficeLocation){
      this.addToColumns(columnToShow,'OfficeLocation');
    }
    if(this.props.isbusinessPhone){
      this.addToColumns(columnToShow,'businessPhone');
    }
    this.state = {
      loading: false,
      selectedKey: "byFirstName",
      columns: columnToShow
    };
  }
  private _handleLinkClick = (item: PivotItem): void => {
    this.setState({
      selectedKey: item.props.itemKey
    });
  }
  private addToColumns(columnToShow : IColumn[],fieldName:string):void{
    const c = columns.filter(x=>x.fieldName === fieldName);
    if(c.length > 0){
      columnToShow.push(c[0]);
    }
  }
  public render(): React.ReactElement<IDelphiTelephonedirProps> {

    return (
      <div className={styles.telephonedirectory} style={{border: this.props.showBorder ? '1px solid' : 'none'}}>
        <div style={{padding:'1%'}}>
          <div>
            <div>
              <WebPartTitle displayMode={this.props.DisplayMode}
                title={this.props.WebpartTitle}
                updateProperty={this.props.updateProperty} />

              <Pivot headersOnly={true}
                selectedKey={this.state.selectedKey}
                onLinkClick={this._handleLinkClick}
                linkSize={PivotLinkSize.normal}
                linkFormat={PivotLinkFormat.tabs}>
                <PivotItem
                  headerText='Search User By First Name'
                  itemKey='byFirstName'
                  itemIcon="Group" />
                <PivotItem
                  headerText='Search User By Last Name'
                  itemKey='byLastName'
                  itemIcon="Group" />
                <PivotItem
                  headerText='Search User By Email'
                  itemKey="byEmail"
                  itemIcon="Group" />
              </Pivot><br />
              {this.state.selectedKey === "byFirstName" &&
                <ByFirstName
                  MSGraphClient={this.props.MsGraphClient}
                  MSGraphServiceInstance={this.props.MSGraphServiceInstance}
                  context={this.props.context}
                  Columns={this.state.columns} />
              }
              {this.state.selectedKey === "byLastName" &&
                <ByLastName
                  MSGraphClient={this.props.MsGraphClient}
                  MSGraphServiceInstance={this.props.MSGraphServiceInstance}
                  context={this.props.context}
                  columns={this.state.columns} />
              }
              {this.state.selectedKey === "byEmail" &&
                <ByEmail
                  MSGraphClient={this.props.MsGraphClient}
                  MSGraphServiceInstance={this.props.MSGraphServiceInstance}
                  context={this.props.context}
                  columns={this.state.columns} />
              }
            </div>
          </div>
        </div>
      </div>
    );
  }
}
