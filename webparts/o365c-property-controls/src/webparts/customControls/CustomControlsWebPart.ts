import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { type IPropertyPaneConfiguration } from "@microsoft/sp-property-pane";
import {
  BaseClientSideWebPart,
  WebPartContext,
} from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";

import * as strings from "CustomControlsWebPartStrings";
import CustomControls from "./components/CustomControls";
import { ICustomControlsProps } from "./components/ICustomControlsProps";
import { ConsoleListener, Logger } from "@pnp/logging";
import WebpartMapper from "../../common/mappers/webpartMapper";
import { SPFx, spfi } from "@pnp/sp";
import SPService from "../../common/services/SPService";
import { IDropdownOption } from "@fluentui/react";
import { IListInfo } from "@pnp/sp/lists/types";

export interface ICustomControlsWebPartProps {
  title?: string;
  description?: string;
  siteUrl: string;
  listId: string;
}

const LOG_SOURCE: string = "CustomControlsWebPart";
export default class CustomControlsWebPart extends BaseClientSideWebPart<ICustomControlsWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = "";
  private propertyPaneConfiguration: (
    properties: ICustomControlsWebPartProps,
    render: () => void,
    context: WebPartContext,
    loadLists: () => Promise<IDropdownOption[]>
  ) => IPropertyPaneConfiguration;

  public render(): void {
    const element: React.ReactElement<ICustomControlsProps> =
      React.createElement(CustomControls, {
        headerProps: WebpartMapper.mapHeader(this.properties),
        listId: this.properties.listId,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
      });

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    // return this._getEnvironmentMessage().then((message) => {
    //   this._environmentMessage = message;
    // });
    this._environmentMessage = await this._getEnvironmentMessage();
    // subscribe a listener
    Logger.subscribe(
      ConsoleListener(LOG_SOURCE, { warning: "#e36c0b", error: "#a80000" })
    );

    //Init SharePoint Service
    const sp = spfi().using(SPFx(this.context));
    SPService.Init(sp);

    return super.onInit();
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app
        .getContext()
        .then((context) => {
          let environmentMessage: string = "";
          switch (context.app.host.name) {
            case "Office": // running in Office
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOffice
                : strings.AppOfficeEnvironment;
              break;
            case "Outlook": // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOutlook
                : strings.AppOutlookEnvironment;
              break;
            case "Teams": // running in Teams
            case "TeamsModern":
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentTeams
                : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(
      this.context.isServedFromLocalhost
        ? strings.AppLocalEnvironmentSharePoint
        : strings.AppSharePointEnvironment
    );
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty(
        "--bodyText",
        semanticColors.bodyText || null
      );
      this.domElement.style.setProperty("--link", semanticColors.link || null);
      this.domElement.style.setProperty(
        "--linkHovered",
        semanticColors.linkHovered || null
      );
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }
  // protected get disableReactivePropertyChanges(): boolean {
  //   return true;
  // }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'custom-controls-property-pane' */
      "./CustomControlsWebPartPropertyPane"
    ).then((importedModule) => {
      this.propertyPaneConfiguration =
        importedModule.getPropertyPaneConfiguration;
    });
  }

  protected onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: any,
    newValue: any
  ): void {
    if (propertyPath === "listId" && oldValue !== newValue) {
      this.properties.listId = newValue;
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this.propertyPaneConfiguration(
      this.properties,
      this.render.bind(this),
      this.context,
      this.loadLists.bind(this)
    );
  }
  private loadLists(): Promise<IDropdownOption[]> {
    return new Promise<IDropdownOption[]>(async (resolve, reject) => {
      let options: IDropdownOption[] = [];

      const lists: IListInfo[] = await SPService.getLists(
        this.properties.siteUrl
      );
      options = [...lists.map((list) => ({ key: list.Id, text: list.Title }))];
      resolve(options);
    });
  }
}
