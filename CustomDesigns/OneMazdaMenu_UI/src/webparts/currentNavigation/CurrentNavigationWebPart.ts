import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { SPComponentLoader } from '@microsoft/sp-loader';
import * as strings from 'CurrentNavigationWebPartStrings';
import CurrentNavigation from './components/CurrentNavigation';
import { ICurrentNavigationProps } from './components/ICurrentNavigationProps';

export interface ICurrentNavigationWebPartProps {
  description: string;
}

export default class CurrentNavigationWebPart extends BaseClientSideWebPart<ICurrentNavigationWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<ICurrentNavigationProps> = React.createElement(
      CurrentNavigation,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);

    this.domElement.innerHTML = `<div id="sideNavBox"><div id="OneMazdaCurrentNavigationUserControl" class="noindex ms-core-listMenu-verticalBox">
    <ul class="root ms-core-listMenu-root static">
      <li class="static">
        <a class="static selected menu-item ms-core-listMenu-item ms-displayInline ms-core-listMenu-selected ms-navedit-linkNode" href="/nametagqa/Pages/DashBoard.aspx">
          <span class="additional-background ms-navedit-flyoutArrow">
            <span class="menu-item-text">Name Tag QA</span>
          </span>
        </a>
      </li>
      <li class="static">
        <a class="menu-item ms-core-listMenu-item ms-displayInline ms-navedit-linkNode selected-sub" href="/nametagqa/Pages/DashBoard.aspx">
          <span class="additional-background ms-navedit-flyoutArrow">
            <span class="menu-item-text">Dashboard</span>
          </span>
        </a>
      </li>
      <li class="static">
        <a class="menu-item ms-core-listMenu-item ms-displayInline ms-navedit-linkNode" href="/nametagqa/Pages/CreateOrder.aspx">
          <span class="additional-background ms-navedit-flyoutArrow">
            <span class="menu-item-text"> Create Order</span>
          </span>
        </a>
      </li>
      <li class="static">
        <a class="menu-item ms-core-listMenu-item ms-displayInline ms-navedit-linkNode" href="/nametagqa/Pages/ProcessedOrder.aspx">
          <span class="additional-background ms-navedit-flyoutArrow">
            <span class="menu-item-text">Processed Order</span>
          </span>
        </a>
      </li>
      <li class="static">
        <a class="menu-item ms-core-listMenu-item ms-displayInline ms-navedit-linkNode" href="/nametagqa/_layouts/15/viewlsts.aspx">
          <span class="additional-background ms-navedit-flyoutArrow">
            <span class="menu-item-text">Site Contents</span>
          </span>
        </a>
      </li>
    </ul>
  </div></div>`
  }

  protected onInit(): Promise<void> {
    SPComponentLoader.loadCss('https://mazdausa.sharepoint.com/sites/MCISPOTestSite/SiteAssets/Branding/css/OneMazda_PubCollab.css');
 
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
