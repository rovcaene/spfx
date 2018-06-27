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

import * as strings from 'RecentUserDocumentsWebPartStrings';
import RecentUserDocuments, { IRecentUserDocumentsProps } from './components/RecentUserDocuments';
import { IUser } from '../people/interfaces';
import { IDynamicDataSource } from '@microsoft/sp-dynamic-data';

export interface IRecentUserDocumentsWebPartProps {
  /**
   * The ID of the dynamic data to which the web part is subscribed
   */
  propertyId: string;
  /**
   * The dynamic data source ID to which the web part is subscribed
   */
  sourceId: string;
  /**
   * Web part title
   */
  title: string;
}

export default class RecentUserDocumentsWebPart extends BaseClientSideWebPart<IRecentUserDocumentsWebPartProps> {
  /**
   * The previous ID of the dynamic data source to which the web part is
   * subscribed. Used to unsubscribe from previously registered dynamic data
   * source notifications after changing web part configuration in the property
   * pane.
   */
  private _lastSourceId: string = undefined;
  /**
   * The previous ID of the dynamic data to which the web part is subscribed.
   * Used to unsubscribe from previously registered dynamic data source
   * notifications after changing web part configuration in the property pane.
   */
  private _lastPropertyId: string = undefined;

  /**
   * Event handler for clicking the Configure button on the Placeholder
   */
  private _onConfigure = (): void => {
    this.context.propertyPane.open();
  }

  protected onInit(): Promise<void> {
    // bind render method to the current instance so that it can be correctly
    // invoked when dynamic data change notification is triggered
    this.render = this.render.bind(this);
    return Promise.resolve();
  }

  public render(): void {
    let user: IUser = undefined;
    const needsConfiguration: boolean = !this.properties.sourceId || !this.properties.propertyId;

    // subscribe to dynamic data changes notifications
    // do this only once the first time the web part is rendered and only,
    // if the dynamic data source ID and property ID are provided
    if (this.renderedOnce === false && !needsConfiguration) {
      try {
        this.context.dynamicDataProvider.registerPropertyChanged(this.properties.sourceId, this.properties.propertyId, this.render);
        // store current values for the dynamic data source ID and property ID
        // so that the web part can unsubscribe from notifications when the
        // web part configuration changes
        this._lastSourceId = this.properties.sourceId;
        this._lastPropertyId = this.properties.propertyId;
      }
      catch (e) {
        this.context.statusRenderer.renderError(this.domElement, `An error has occurred while connecting to the data source. Details: ${e}`);
        return;
      }
    }

    // retrieve the current value of dynamic data only if the dynamic data
    // source ID and property ID have been provided
    if (!needsConfiguration) {
      const source: IDynamicDataSource = this.context.dynamicDataProvider.tryGetSource(this.properties.sourceId);
      user = source ? source.getPropertyValue(this.properties.propertyId) : undefined;
    }

    const element: React.ReactElement<IRecentUserDocumentsProps> = React.createElement(
      RecentUserDocuments,
      {
        context: this.context,
        user: user
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
    // get all available dynamic data sources on the page
    const sourceOptions: IPropertyPaneDropdownOption[] =
      this.context.dynamicDataProvider.getAvailableSources().map(source => {
        return {
          key: source.id,
          text: source.metadata.title
        };
      });
    const selectedSource: string = this.properties.sourceId;

    let propertyOptions: IPropertyPaneDropdownOption[] = [];
    if (selectedSource) {
      const source: IDynamicDataSource = this.context.dynamicDataProvider.tryGetSource(selectedSource);
      if (source) {
        // get the list of all properties exposed by the currently selected
        // data source
        propertyOptions = source.getPropertyDefinitions().map(prop => {
          return {
            key: prop.id,
            text: prop.title
          };
        });
      }
    }

    return {
      pages: [
        {
          groups: [
            {
              groupFields: [
                PropertyPaneDropdown('sourceId', {
                  label: strings.SourceIdFieldLabel,
                  options: sourceOptions,
                  selectedKey: this.properties.sourceId
                }),
                PropertyPaneDropdown('propertyId', {
                  label: strings.PropertyIdFieldLabel,
                  options: propertyOptions,
                  selectedKey: this.properties.propertyId
                })
              ]
            }
          ]
        }
      ]
    };
  }

  protected onPropertyPaneFieldChanged(propertyPath: string): void {
    if (propertyPath === 'sourceId') {
      // reset the selected property ID after selecting a different dynamic
      // data source
      this.properties.propertyId =
        this.context.dynamicDataProvider.tryGetSource(this.properties.sourceId).getPropertyDefinitions()[0].id;
    }

    if (this._lastSourceId && this._lastPropertyId) {
      // unsubscribe from the previously registered dynamic data changes
      // notifications
      this.context.dynamicDataProvider.unregisterPropertyChanged(this._lastSourceId, this._lastPropertyId, this.render);
    }

    // subscribe to the newly configured dynamic data changes notifications
    this.context.dynamicDataProvider.registerPropertyChanged(this.properties.sourceId, this.properties.propertyId, this.render);
    // store current values for the dynamic data source ID and property ID
    // so that the web part can unsubscribe from notifications when the
    // web part configuration changes
    this._lastSourceId = this.properties.sourceId;
    this._lastPropertyId = this.properties.propertyId;
  }
}
