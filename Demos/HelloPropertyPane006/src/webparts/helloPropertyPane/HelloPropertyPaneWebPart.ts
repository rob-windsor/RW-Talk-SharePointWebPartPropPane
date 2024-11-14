import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'HelloPropertyPaneWebPartStrings';
import HelloPropertyPane from './components/HelloPropertyPane';
import { IHelloPropertyPaneProps } from './components/IHelloPropertyPaneProps';

import { SPHttpClient } from '@microsoft/sp-http';
import { PropertyFieldPeoplePicker, PrincipalType, IPropertyFieldGroupOrPerson } from '@pnp/spfx-property-controls';

export interface IList {
  Id: string;
  Title: string;
}

export interface IListCollection {
  value: IList[];
}

export interface IHelloPropertyPaneWebPartProps {
  description: string;
  color: string;
  list: string;
  users: IPropertyFieldGroupOrPerson[];
}

export default class HelloPropertyPaneWebPart extends BaseClientSideWebPart<IHelloPropertyPaneWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  private lists: IPropertyPaneDropdownOption[] | undefined = undefined;
  private listsDropdownDisabled: boolean = true;  

  private getUserName(): string {
    let result = "";

    if (this.properties.users && this.properties.users.length == 1) {
      result = this.properties.users[0].fullName;
    }

    return result;
  }

  public render(): void {
    const element: React.ReactElement<IHelloPropertyPaneProps> = React.createElement(
      HelloPropertyPane,
      {
        description: this.properties.description,
        color: this.properties.color,
        list: this.properties.list,
        user: this.getUserName(),
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
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

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  private validateDescription(value: string): string {
    let result = "";

    if (value == null || value.trim().length === 0) {
      result = "Please enter a description";
    }

    return result;
  } 

  private async loadLists(): Promise<IPropertyPaneDropdownOption[]> {
    let restUrl = this.context.pageContext.web.absoluteUrl +
      "/_api/web/lists?$filter=(Hidden eq false)";

    let result: IPropertyPaneDropdownOption[] = [];
    try {
      const response = await this.context.spHttpClient.get(restUrl, SPHttpClient.configurations.v1);
      if (response.ok) {
        const data = await response.json();
        result = data.value.map((list: any) => {
          return {key: list.Id, text: list.Title};
        });
      }
    } catch (ex) {
      console.log(ex);
    }

    return result;
  }

  protected onPropertyPaneConfigurationStart(): void {
    if (this.lists) {
      return;
    }

    this.listsDropdownDisabled = true;

    this.loadLists()
      .then((listOptions: IPropertyPaneDropdownOption[]): void => {
        this.lists = listOptions;
        this.listsDropdownDisabled = false;
        this.context.propertyPane.refresh();
        this.render();
      });
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
                  label: strings.DescriptionFieldLabel,
                  onGetErrorMessage: this.validateDescription.bind(this)
                }),
                PropertyPaneDropdown("color", {
                  label: "Color",
                  options: [
                    { key: "Red", text: "Red" },
                    { key: "Green", text: "Green" },
                    { key: "Blue", text: "Blue" }
                  ]
                }),
                PropertyPaneDropdown("list", {
                  label: "List",
                  options: this.lists,
                  disabled: this.listsDropdownDisabled
                }),
                PropertyFieldPeoplePicker('users', {
                  label: 'User',
                  initialData: this.properties.users,
                  allowDuplicate: false,
                  multiSelect: false,
                  principalType: [PrincipalType.Users, PrincipalType.SharePoint, PrincipalType.Security],
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  context: this.context as any,
                  properties: this.properties,
                  deferredValidationTime: 0,
                  key: 'peopleFieldId'
                })                            
              ]
            }
          ]
        }
      ]
    };
  }
}
