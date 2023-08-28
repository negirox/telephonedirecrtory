import * as React from "react";
import { ByEmailProps } from "./ByEmailProps";
import styles from '../DelphiTelephonedir.module.scss';
import { ByEmailState } from "./ByEmailState";
//import { autobind } from "office-ui-fabric-react/lib/Utilities";
import { PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { Stack, IStackProps, IStackStyles } from 'office-ui-fabric-react/lib/Stack';
import { DetailsList, DetailsListLayoutMode } from 'office-ui-fabric-react/lib/DetailsList';
import { Log } from "@microsoft/sp-core-library";
const stackTokens = { childrenGap: 50 };
const LOG_SOURCE = "ByEmail";
const stackStyles: Partial<IStackStyles> = { root: { width: 650 } };
const columnProps: Partial<IStackProps> = {
  tokens: { childrenGap: 15 },
  styles: { root: { width: 700 } },
};

export class ByEmail extends React.Component<ByEmailProps, ByEmailState>{
  constructor(props: ByEmailProps) {
    super(props);
    this.state = {
      loading: false,
      searchFor: '',
      userProperties: [],
      isDataFound: true,
    };
    this.getUsers = this.getUsers.bind(this);
    this._getPeoplePickerItems = this._getPeoplePickerItems.bind(this);
  }
  private async _getPeoplePickerItems(items: any[]):Promise<void> {  
    console.log('Items:', items);  
    if(items.length > 0){
      try {
        const searchItem = items[0].id;
        const splitsText = searchItem.split('|');
        const searchFor = splitsText[splitsText.length -1];
        this.setState({
          searchFor: searchFor,
        });
        await this.getUsers(searchFor);
      } catch (error) {
        Log.error(LOG_SOURCE, error);
      }
    }
    else{
      this.setState({
        searchFor: '',
        userProperties : [],
        isDataFound: false
      });
    }
  } 


  //@autobind
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  public async getUsers(email: string): Promise<any> {
    this.setState({ loading: true }, async () => {
      await this.props.MSGraphServiceInstance
        .getUserProperties(email, this.props.MSGraphClient)
        // tslint:disable-next-line: no-shadowed-variable
        .then((users) => {
          if (users.length !== 0) {
            this.setState({
              userProperties: users,
              isDataFound: true
            });
          }
          else {
            this.setState({
              userProperties: [],
              isDataFound: false
            });
          }
        });
    });
  }
  public render(): React.ReactElement<ByEmailProps> {
    return (
      <div className={styles.telephonedirectory}>
        <div>
          <Stack horizontal tokens={stackTokens} styles={stackStyles}>
            <Stack {...columnProps}>
              <PeoplePicker
                context={this.props.context}
                placeholder=""
                titleText="Email"
                personSelectionLimit={1}
                showtooltip={true}
                disabled={false}
                principalTypes={[PrincipalType.User]}
                resolveDelay={1000} 
                selectedItems={this._getPeoplePickerItems}   
              />
            </Stack>
          </Stack>
          <div />
          <div id='detailedList'>
            {this.state.userProperties.length !== 0 &&
              <DetailsList
                items={this.state.userProperties}
                columns={this.props.columns}
                isHeaderVisible={true}
                layoutMode={DetailsListLayoutMode.justified}
                usePageCache={true}
              />
            }
          </div>
        </div>
      </div>
    );
  }
}
