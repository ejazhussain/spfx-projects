import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as strings from 'UploadLargeFileWebPartStrings';
import { MSGraphClientV3 } from "@microsoft/sp-http";
import GraphService from '../services/GraphService';
import { UploadFile } from './components/UploadFile';
import { IUploadFileProps } from './components/IUploadFile';
import { initializeIcons } from '@fluentui/font-icons-mdl2';
import { Logger, LogLevel, ConsoleListener } from "@pnp/logging";

export interface IUploadLargeFileWebPartProps {
  description: string;
}
const LOG_SOURCE: string = 'UPLOADINGLARGEFILE';
export default class UploadLargeFileWebPart extends BaseClientSideWebPart<IUploadLargeFileWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  private _graphClient:MSGraphClientV3;
  



  public render(): void {
    const element: React.ReactElement<IUploadFileProps> = React.createElement(
      UploadFile,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        graphClient: this._graphClient
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    //Logging
    // subscribe a listener
    Logger.subscribe(ConsoleListener(LOG_SOURCE, {warning:'#e36c0b',error:'#a80000', info:'#881798'}));
    // Logger.subscribe(ConsoleListener());
    // set the active log level -- eventually make this a web part property
    Logger.activeLogLevel = LogLevel.Info;

    initializeIcons();
    // Create the Microsoft Graph client
    this._graphClient = await  this.context.msGraphClientFactory.getClient('3');     

    GraphService.init(this._graphClient);
    
    return super.onInit();
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
