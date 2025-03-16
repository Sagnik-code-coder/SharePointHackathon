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
import * as strings from 'PageFooterWebPartStrings';
import PageFooter from './components/PageFooter';
import { IPageFooterProps } from './components/IPageFooterProps';

export interface IPageFooterWebPartProps {
  description: string;
}

export default class PageFooterWebPart extends BaseClientSideWebPart<IPageFooterWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IPageFooterProps> = React.createElement(
      PageFooter,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);

    this.domElement.innerHTML = `<div class="footer ms-dialogHidden hidden-print">
    <div class="footer-container">
      <div class="top-footer-wrapper">
                    <div class="top-footer">
                        <div class="customNav">
                <table align="center" border="0">
                <tbody><tr>
                <td valign="top" style="width:250px" class="
                menu-level2">
                <table style="padding:0px;border-spacing:0px;margin-left:50px;">
                <tbody><tr>
                <td>
                <a href="#" description="IT Support">IT Support</a>
                </td>
                </tr>
                </tbody></table>
                <table style="padding:0px;border-spacing:0px;margin-left:50px;">
                <tbody><tr>
                <td>
                <a href="#" description="Parts Support">Parts Support</a>
                </td>
                </tr>
                </tbody></table>
                <table style="padding:0px;border-spacing:0px;margin-left:50px;">
                <tbody><tr>
                <td>
                <a href="#" description="Technical Support">Technical Support</a>
                </td>
                </tr>
                </tbody></table>
                <table style="padding:0px;border-spacing:0px;margin-left:50px;">
                <tbody><tr>
                <td>
                <a href="#" description="#">Warranty Support</a>
                </td>
                </tr>
                </tbody></table>
                </td>
                <td valign="top" style="width:250px" class="
                menu-level2">
                <table style="padding:0px;border-spacing:0px;margin-left:50px;">
                <tbody><tr>
                <td>
                <a href="#" description="Organizational Chart">Organizational Chart</a>
                </td>
                </tr>
                </tbody></table>
                </td>
                <td valign="top" style="width:250px" class="
                menu-level2">
                </td><td valign="top" style="width:250px" class="
                menu-level2">
                </td><td valign="top" style="width:250px" class="
                menu-level2">
                </td></tr>
                </tbody></table>
                </div>
                    </div>
                </div>
                <div class="site-navigation">
                    <div class="site-navigation-text"></div>
                    <div class="ms-breadcrumb-dropdownBox" style="">
                        <span id="DeltaBreadcrumbDropdown">
                            <span class="ms-breadcrumb-anchor"><span class="s4-clust" style="height:16;width:16;position:relative;display:inline-block;overflow:hidden;"><a id="GlobalBreadCrumbNavPopout-anchor" onclick="CoreInvoke('callOpenBreadcrumbMenu', event, 'GlobalBreadCrumbNavPopout-anchor', 'GlobalBreadCrumbNavPopout-menu', 'GlobalBreadCrumbNavPopout-img', 'ms-breadcrumb-anchor-open', 'ltr', '', false); return false;" onmouseover="" onmouseout="" title="Navigate Up" href="javascript:;" style="display:inline-block;height:16px;width:16px;"><img src="/_layouts/15/images/spcommon.png?rev=23" alt="Navigate Up" style="border-width:0;position:absolute;left:-215px !important;top:-120px !important;"></a></span></span><div class="ms-popoutMenu ms-breadcrumb-menu ms-noList" id="GlobalBreadCrumbNavPopout-menu" style="display:none;"><div class="ms-breadcrumb-top">
                                    <span class="ms-breadcrumb-header">This page location is:</span></div><ul class="ms-breadcrumb">
                <li class="ms-breadcrumbRootNode"><span class="s4-breadcrumb-arrowcont"><span style="height:16px;width:16px;position:relative;display:inline-block;overflow:hidden;" class="s4-clust s4-breadcrumb"><img src="/_layouts/15/images/spcommon.png?rev=23" alt="" style="position:absolute;left:-217px !important;top:-210px !important;"></span></span><a title="Home" class="ms-breadcrumbRootNode" href="/">Home</a><ul class="ms-breadcrumbRootNode"><li class="ms-breadcrumbNode"><span class="s4-breadcrumb-arrowcont"><span style="height:16px;width:16px;position:relative;display:inline-block;overflow:hidden;" class="s4-clust s4-breadcrumb"><img src="/_layouts/15/images/spcommon.png?rev=23" alt="" style="position:absolute;left:-217px !important;top:-210px !important;"></span></span><a title="Pages" class="ms-breadcrumbNode" href="https://stg-one.mazda.ca/_layouts/15/listform.aspx?ListId=%7B2EE78DF5%2D67DF%2D466B%2D8F68%2D574B32FA9831%7D&amp;PageType=0">Pages</a><ul class="ms-breadcrumbNode"><li class="ms-breadcrumbCurrentNode"><span class="s4-breadcrumb-arrowcont"><span style="height:16px;width:16px;position:relative;display:inline-block;overflow:hidden;" class="s4-clust s4-breadcrumb"><img src="/_layouts/15/images/spcommon.png?rev=23" alt="" style="position:absolute;left:-217px !important;top:-210px !important;"></span></span><span class="ms-breadcrumbCurrentNode">default</span></li></ul></li></ul></li>
                </ul></div>
                        </span>
                    </div>
                </div>

                <div class="bottom-footer-wrapper">
                    <div class="bottom-footer">
                        <div class="footerrow">
                            <div class="span-12">
                                <nav class="footer-links">
                                    <ul class="clearfix footer-ul">
                                      
                                        <li>
                                            <!--<pointfire:control sectionid="Navigationfrench" language="1036" method="hide" />--><span style="display:none;"><span style="display:none;"><a href="mailto:webteam@mazda.ca?subject=Commentaires One.Mazda">Envoyer vos commentaires</a></span></span><!--<pointfire:control endsection="Navigationfrench" />-->
                                <!--<pointfire:control sectionid="Navigationenglish" language="1033" method="hide" />--><a href="mailto:webteam@mazda.ca?subject=One.Mazda Feedback">Website Feedback</a><!--<pointfire:control endsection="Navigationenglish" />--> 
                              </li>
                                    </ul>
                                </nav>
                                <nav class="footer-social-links">
                                    <ul class="clearfix">
                                        <li>
                                        <!--<pointfire:control sectionid="Navigationfrench" language="1036" method="hide" />--><span style="display:none;"><span style="display:none;">	<span>Suivez-nous:</span></span></span><!--<pointfire:control endsection="Navigationfrench" />-->
                              <!--<pointfire:control sectionid="Navigationenglish" language="1033" method="hide" />-->	<span>Follow us:</span><!--<pointfire:control endsection="Navigationenglish" />--> 
                              </li>
                                        <li><a href="#" target="_blank" class="footer-facebook"><span class="hide-link-text">Facebook</span></a></li>
                                        <li><a href="#" target="_blank" class="footer-twitter"><span class="hide-link-text">Twitter</span></a></li>
                                        <!--<li><a href="#" target="_blank" class="footer-pinterest"><span class="hide-link-text">Pinterest</span></a></li>-->
                                        <li><a href="#" target="_blank" class="footer-instagram"><span class="hide-link-text">Instagram</span></a></li>
                                    </ul>
                                </nav>
                            </div>
                        </div>
                    </div>
                </div>
              </div>
            </div>`;
  }

  protected onInit(): Promise<void> {
    SPComponentLoader.loadCss('https://mazdausa.sharepoint.com/sites/MCISPOTestSite/SiteAssets/Branding/css/OneMazda_PubCollab.css');
    //SPComponentLoader.loadCss('https://mazdausa.sharepoint.com/sites/MCISPOTestSite/SiteAssets/Branding/css/AppLinks.css');
    //SPComponentLoader.loadCss('https://mazdausa.sharepoint.com/sites/MCISPOTestSite/SiteAssets/NameTag/css/bootstrap.min.css');    
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
