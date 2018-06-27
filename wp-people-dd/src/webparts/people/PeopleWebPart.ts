import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown
} from '@microsoft/sp-webpart-base';

import * as strings from 'PeopleWebPartStrings';
import People, { IPeopleProps } from './components/People';
import { IDataService, ISPMemberInfo, IUser } from './interfaces/index';
import { DataService } from './dal/DataService';
import { IDynamicDataController, IDynamicDataPropertyDefinition } from '@microsoft/sp-dynamic-data';

export interface IPeopleWebPartProps {
  selectedGroup?: number;
}

export default class PeopleWebPart extends BaseClientSideWebPart<IPeopleWebPartProps> implements IDynamicDataController {
  private _memberOptions: IPropertyPaneDropdownOption[] = [];
  private _dataService: IDataService;
  private _selectedUser: IUser;

  constructor() {
    super();
    this.onUserSelected = this.onUserSelected.bind(this);
  }

  /**
     * Returns all the property definitions for dynamic data.
     * This needs to be overriden by the implementation of the component.
     */
  public getPropertyDefinitions(): ReadonlyArray<IDynamicDataPropertyDefinition> {
    return [
      {
        id: 'user',
        title: 'User'
      }
    ];
  }
  /**
   * Given a property id, returns the value of the property.
   * This needs to be overriden by the implementation of the component.
   */
  public getPropertyValue(propertyId: string): any {
    switch (propertyId) {
      case 'user':
        return this._selectedUser;
    }

    throw new Error('Bad property id');
  }

  protected onInit(): Promise<void> {
    // Initialize the DynamicDataSource Manager
    this.context.dynamicDataSourceManager.initializeSource(this);

    // DAL service
    const dataService = new DataService(this.context);

    // Get SP members
    dataService.GetMemberInfo().then((members: ISPMemberInfo[]) => {
      this._memberOptions = members.map((mi: ISPMemberInfo) => {
        return {
          key: mi.Id,
          text: mi.Title
        };
      });
    });

    return Promise.resolve();
  }

  protected onUserSelected(user: IUser) {
    this._selectedUser = user;
    if(this._selectedUser)
      console.log(`Selected ${this._selectedUser}`);

    this.context.dynamicDataSourceManager.notifyPropertyChanged('user');
  }

  public render(): void {
    const element: React.ReactElement<IPeopleProps> = React.createElement(
      People,
      {
        context: this.context,
        selectedGroup: this.properties.selectedGroup,
        onUserSelected: this.onUserSelected
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
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneDropdown('selectedGroup', {
                  label: strings.SelectedGroupFieldLabel,
                  options: this._memberOptions
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
