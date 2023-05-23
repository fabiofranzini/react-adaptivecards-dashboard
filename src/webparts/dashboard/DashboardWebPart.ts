import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'DashboardWebPartStrings';
import Dashboard, { IDashboardProps } from './components/Dashboard';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { IColumnReturnProperty, PropertyFieldColumnPicker, PropertyFieldColumnPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldColumnPicker';

export interface IDashboardWebPartProps {
  title: string;
  widgetIds: string[];
  canteenOrderList: string;
  emailColumn: string;
  entreeColumn: string;
  sideColumn: string;
  drinkColumn: string;
}

export default class DashboardWebPart extends BaseClientSideWebPart<IDashboardWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IDashboardProps> = React.createElement(
      Dashboard,
      {
        context: this.context,
        widgetIds: this.properties.widgetIds,
        canteenOrderList: this.properties.canteenOrderList,
        emailColumn: this.properties.emailColumn,
        entreeColumn: this.properties.entreeColumn,
        sideColumn: this.properties.sideColumn,
        drinkColumn: this.properties.drinkColumn,
        title: this.properties.title,
        onAddWidgets: (widgetIds: string[]) => this.properties.widgetIds = widgetIds,
        onUpdateTitle: (title: string) => this.properties.title = title,
        displayMode: this.displayMode
      }
    );

    ReactDom.render(element, this.domElement);
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
            description: null
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: 'Dashboard Title'
                })
              ]
            },
            {
              groupName: "Canteen Order Widget Settings",
              groupFields: [
                PropertyFieldListPicker('canteenOrderList', {
                  label: 'Select a list for Canteen Orders',
                  selectedList: this.properties.canteenOrderList,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context as any,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'canteenOrderListField'
                }),
                PropertyFieldColumnPicker('emailColumn', {
                  label: "Select the 'email' column",
                  context: this.context as any,
                  selectedColumn: this.properties.emailColumn,
                  listId: this.properties.canteenOrderList,
                  disabled: false,
                  orderBy: PropertyFieldColumnPickerOrderBy.Title,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'emailColumnFieldId',
                  displayHiddenColumns: false,
                  columnReturnProperty: IColumnReturnProperty["Internal Name"]
                }),
                PropertyFieldColumnPicker('entreeColumn', {
                  label: "Select the 'entree' column",
                  context: this.context as any,
                  selectedColumn: this.properties.entreeColumn,
                  listId: this.properties.canteenOrderList,
                  disabled: false,
                  orderBy: PropertyFieldColumnPickerOrderBy.Title,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'entreeColumnFieldId',
                  displayHiddenColumns: false,
                  columnReturnProperty: IColumnReturnProperty["Internal Name"]
                }),
                PropertyFieldColumnPicker('sideColumn', {
                  label: "Select the 'side' column",
                  context: this.context as any,
                  selectedColumn: this.properties.sideColumn,
                  listId: this.properties.canteenOrderList,
                  disabled: false,
                  orderBy: PropertyFieldColumnPickerOrderBy.Title,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'sideColumnFieldId',
                  displayHiddenColumns: false,
                  columnReturnProperty: IColumnReturnProperty["Internal Name"]
                }),
                PropertyFieldColumnPicker('drinkColumn', {
                  label: "Select the 'drink' column",
                  context: this.context as any,
                  selectedColumn: this.properties.drinkColumn,
                  listId: this.properties.canteenOrderList,
                  disabled: false,
                  orderBy: PropertyFieldColumnPickerOrderBy.Title,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'drinkColumnFieldId',
                  displayHiddenColumns: false,
                  columnReturnProperty: IColumnReturnProperty["Internal Name"]
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
