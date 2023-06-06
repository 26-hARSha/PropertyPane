import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,

  PropertyPaneCheckbox,

  PropertyPaneChoiceGroup,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'PropertyPaneWebPartStrings';
import PropertyPane from './components/PropertyPane';
import { IPropertyPaneProps } from './components/IPropertyPaneProps';
//import { Label } from 'office-ui-fabric-react';

export interface IPropertyPaneWebPartProps {
  Discount: boolean;
  ChoiceGroups: string;
  DropDown: any;
  IsMarried: boolean;
  Hobbie: any;
  description: string;
  getUserName: string;
  getAge: Number;

}

export default class PropertyPaneWebPart extends BaseClientSideWebPart<IPropertyPaneWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IPropertyPaneProps> = React.createElement(
      PropertyPane,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        siteAbsoluteURL: this.context.pageContext.web.absoluteUrl,
        siteTitle: this.context.pageContext.web.title,
        //Property Pane
        getUserName: this.properties.getUserName,
        getAge: this.properties.getAge,
        Hobbie: this.properties.Hobbie,
        IsMarried: this.properties.IsMarried,
        DropDown: this.properties.DropDown,
        ChoiceGroups: this.properties.ChoiceGroups,
        Discount: this.properties.Discount
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
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error('Unknown host');
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
  //For Apply Button 
  protected get disableReactivePropertyChanges(): boolean { return true; }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {

          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("getUserName", { label: "Enter Your Full Name", }),
                PropertyPaneSlider("getAge", {
                  label: "Select Your Age",
                  min: 18,
                  max: 99,
                }),
                PropertyPaneChoiceGroup("Hobbie", {
                  label: "Select Your Hobbie",
                  options: [
                    {
                      key: "1",
                      text: "reading",
                    },
                    {
                      key: "2",
                      text: "cooking",
                    },
                    {
                      key: "3",
                      text: "writing",
                    },
                    {
                      key: "4",
                      text: "Drawing",
                    },
                    {
                      key: "5",
                      text: "traveling",
                    },

                  ],
                }),
                PropertyPaneToggle("IsMarried", {
                  label: "Is Married",
                  onText: "Yes", offText: "No",
                }),
                PropertyPaneDropdown("DropDown", {
                  label: "Department",
                  options: [
                    {
                      key: "IT Department",
                      text: "IT Department",
                    },
                    {
                      key: "R&D",
                      text: "R&D",
                    },
                    {
                      key: "HR",
                      text: "HR",
                    },
                    {
                      key: "Finance",
                      text: "Finance",
                    },

                  ], selectedKey: "IT Department"
                }),
                PropertyPaneCheckbox("Discount", {
                  text: "Do You Have a Discount Coupon?",
                  checked: false,
                  disabled: false
                })
              ],

            }
          ]
        }
      ]
      
    };
  }
}
