import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { MSGraphClientV3 } from "@microsoft/sp-http";
import * as strings from 'DelphiTelephonedirWebPartStrings';
import DelphiTelephonedir from './components/DelphiTelephonedir';
import { IDelphiTelephonedirProps } from './components/IDelphiTelephonedirProps';
import { MSGraphService } from '../../Services/MSGraphService';

export interface IDelphiTelephonedirWebPartProps {
  description: string;
  Title: string;
  isDisplayName:boolean,
  isEmail:boolean,
  ismobilePhone:boolean,
  isJobTitle:boolean,
  isOfficeLocation:boolean,
  isbusinessPhone:boolean,
  showBorder:boolean,
}

export default class DelphiTelephonedirWebPart extends BaseClientSideWebPart<IDelphiTelephonedirWebPartProps> {
  private MSGraphServiceInstance: MSGraphService;
  private MSGraphClient: MSGraphClientV3

  public render(): void {
    const element: React.ReactElement<IDelphiTelephonedirProps> = React.createElement(
      DelphiTelephonedir,
      {
        description: this.properties.description,
        MSGraphServiceInstance: this.MSGraphServiceInstance,
        context: this.context,
        MsGraphClient: this.MSGraphClient,
        DisplayMode: this.displayMode,
        WebpartTitle: this.properties.Title,
        updateProperty: (value: string) => {
          this.properties.Title = value;
        },
        isDisplayName:this.properties.isDisplayName ?? true,
        isEmail:this.properties.isEmail ?? true,
        ismobilePhone:this.properties.ismobilePhone ?? true,
        isJobTitle:this.properties.isJobTitle,
        isOfficeLocation:this.properties.isOfficeLocation,
        isbusinessPhone:this.properties.isbusinessPhone,
        showBorder:this.properties.showBorder
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit():Promise<void> {
    await super.onInit();
    this.MSGraphServiceInstance = new MSGraphService();
    this.MSGraphClient = await this.context.msGraphClientFactory.getClient('3');
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get _dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneTextField('Title', {
                  label: strings.Title
                }),
                PropertyPaneToggle('isDisplayName',{
                  label:"Show First Name",
                  onText:"Yes",
                  offText:"No"
                }),
                PropertyPaneToggle('isEmail',{
                  label:"Show Email",
                  onText:"Yes",
                  offText:"No"
                }),
                PropertyPaneToggle('ismobilePhone',{
                  label:"Show Mobile Number",
                  onText:"Yes",
                  offText:"No"
                }),
                PropertyPaneToggle('isJobTitle',{
                  label:"Show Job Title",
                  onText:"Yes",
                  offText:"No"
                }),
                PropertyPaneToggle('isOfficeLocation',{
                  label:"Show Office Location",
                  onText:"Yes",
                  offText:"No"
                }),
                PropertyPaneToggle('isbusinessPhone',{
                  label:"Show Business Phone",
                  onText:"Yes",
                  offText:"No"
                }),
                PropertyPaneToggle('showBorder',{
                  label:"Show Border",
                  onText:"Yes",
                  offText:"No"
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
