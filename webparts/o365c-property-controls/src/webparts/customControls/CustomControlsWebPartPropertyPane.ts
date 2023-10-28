import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ICustomControlsWebPartProps } from "./CustomControlsWebPart";
import {
  IPropertyPaneConfiguration,
  IPropertyPaneField,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { PropertyFieldTextWithCallout } from "@pnp/spfx-property-controls";
import { CalloutTriggers } from "@pnp/spfx-property-controls/lib/common/callout/Callout";
import * as React from "react";
import { PropertyPaneAsyncDropdown } from "../../common/controls/PropertyPaneAsyncDropdown/PropertyPaneAsyncDropdown";
import { IDropdownOption } from "@fluentui/react";
// import { IListInfo } from "@pnp/sp/lists/types";
// import SPService from "../../common/services/SPService";

export function getPropertyPaneConfiguration(
  properties: ICustomControlsWebPartProps,
  render: () => void,
  context: WebPartContext,
  loadLists: () => Promise<IDropdownOption[]>
): IPropertyPaneConfiguration {
  return {
    pages: [
      {
        displayGroupsAsAccordion: true,
        groups: [
          {
            groupName: "Header settings",
            groupFields: getHeaderSettingsFields(),
            isCollapsed: true,
          },
          {
            groupName: "Controls settings",
            groupFields: getControlSettingsFields(),
            isCollapsed: true,
          },
        ].map((f) => f),
      },
    ],
  };

  function getHeaderSettingsFields(): IPropertyPaneField<any>[] {
    const fields: IPropertyPaneField<any>[] = [
      PropertyPaneTextField("title", {
        label: "Title",
      }),
      PropertyPaneTextField("description", {
        label: "Description",
      }),
    ];
    return fields;
  }
  function getControlSettingsFields(): IPropertyPaneField<any>[] {
    const fields: IPropertyPaneField<any>[] = [
      PropertyFieldTextWithCallout("siteUrl", {
        calloutTrigger: CalloutTriggers.Click,
        key: "siteUrlFieldId",
        label: "Site URL",
        calloutContent: React.createElement(
          "span",
          {},
          "URL of the site where the document library to show documents from is located. Leave empty to connect to a document library from the current site"
        ),
        calloutWidth: 250,
        value: properties.siteUrl,
      }),
      new PropertyPaneAsyncDropdown("list", {
        label: "Select a list",
        loadOptions: loadLists,
        selectedKey: properties.listId,
        disabled: !properties.siteUrl,
        onPropertyChange: (propertyPath, newValue) => {
          properties.listId = newValue;
          render();
          context.propertyPane.refresh();
        },
      }),
    ];
    return fields;
  }
}
