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
import * as strings from 'LandingPageWebPartStrings';
import LandingPage from './components/LandingPage';
import { ILandingPageProps } from './components/ILandingPageProps';

export interface ILandingPageWebPartProps {
  description: string;
}

export default class LandingPageWebPart extends BaseClientSideWebPart<ILandingPageWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<ILandingPageProps> = React.createElement(
      LandingPage,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
    this.domElement.innerHTML = `
    <div class="flexslider">      
<div class="flex-viewport" style="overflow: hidden; position: relative;"><ul class="slides" style="width: 1000%; transition-duration: 0.6s; transform: translate3d(-2700px, 0px, 0px);"><li class="clone" aria-hidden="true" style="width: 900px; float: left; display: block;">
                <div class="sliderwrapper">
                    <a href="#">
                        <img class="slideShowImages" src="/sites/MCISPOTestSite/SiteAssets/Branding/images/LandingPageSlider/CarImage1.png" alt="missing" draggable="false">
                    </a>
                    <div class="container">
                        <div class="content">
                            <h2>test</h2>
                            <div>
                                <p></p>
                            </div>
                            <p><a href="https://one.mazda.ca/Pages/eCelebration.aspx" class="use-sprite">
                                <!--<pointfire:control sectionid="Navigationfrench" language="1036" method="hide" />--><span style="display:none;"><span style="display:none;">En savoir plus</span></span><!--<pointfire:control endsection="Navigationfrench" />-->
                           		<!--<pointfire:control sectionid="Navigationenglish" language="1033" method="hide" />-->Find out more<!--<pointfire:control endsection="Navigationenglish" />--> 
                            <span class="sprite-link-arrow">  </span></a></p>
                        </div>
                    </div>
                </div>
            </li>
                
            <li style="width: 900px; float: left; display: block;" class="">
                <div class="sliderwrapper">
                    <a href="http://google.ca">
                        <img class="slideShowImages" src="/sites/MCISPOTestSite/SiteAssets/Branding/images/LandingPageSlider/CarImage2.png" alt="missing" draggable="false">
                    </a>
                    <div class="container">
                        <div class="content">
                            <h2>TEst</h2>
                            <div>
                                <p></p>
                            </div>
                            <p><a href="http://google.ca" class="use-sprite">
                                <!--<pointfire:control sectionid="Navigationfrench" language="1036" method="hide" />--><span style="display:none;"><span style="display:none;">En savoir plus</span></span><!--<pointfire:control endsection="Navigationfrench" />-->
                           		<!--<pointfire:control sectionid="Navigationenglish" language="1033" method="hide" />-->Find out more<!--<pointfire:control endsection="Navigationenglish" />--> 
                            <span class="sprite-link-arrow">  </span></a></p>
                        </div>
                    </div>
                </div>
            </li>
        
            <li class="" style="width: 900px; float: left; display: block;">
                <div class="sliderwrapper">
                    <a href="#" target="_blank">
                        <img class="slideShowImages" src="/sites/MCISPOTestSite/SiteAssets/Branding/images/LandingPageSlider/CarImage3.png" alt="missing" draggable="false">
                    </a>
                    <div class="container">
                        <div class="content">
                            <h2>Test</h2>
                            <div>
                                <p></p>
                            </div>
                            <p><a href="#" class="use-sprite" target="_blank">
                                <!--<pointfire:control sectionid="Navigationfrench" language="1036" method="hide" />--><span style="display:none;"><span style="display:none;">En savoir plus</span></span><!--<pointfire:control endsection="Navigationfrench" />-->
                           		<!--<pointfire:control sectionid="Navigationenglish" language="1033" method="hide" />-->Find out more<!--<pointfire:control endsection="Navigationenglish" />--> 
                            <span class="sprite-link-arrow">  </span></a></p>
                        </div>
                    </div>
                </div>
            </li>
        
            <li class="flex-active-slide" style="width: 900px; float: left; display: block;">
                <div class="sliderwrapper">
                    <a href="#">
                        <img class="slideShowImages" src="/sites/MCISPOTestSite/SiteAssets/Branding/images/LandingPageSlider/CarImage4.png" alt="missing" draggable="false">
                    </a>
                    <div class="container">
                        <div class="content">
                            <h2>test</h2>
                            <div>
                                <p></p>
                            </div>
                            <p><a href="#" class="use-sprite">
                                <!--<pointfire:control sectionid="Navigationfrench" language="1036" method="hide" />--><span style="display:none;"><span style="display:none;">En savoir plus</span></span><!--<pointfire:control endsection="Navigationfrench" />-->
                           		<!--<pointfire:control sectionid="Navigationenglish" language="1033" method="hide" />-->Find out more<!--<pointfire:control endsection="Navigationenglish" />--> 
                            <span class="sprite-link-arrow">  </span></a></p>
                        </div>
                    </div>
                </div>
            </li>
        
            <li style="width: 900px; float: left; display: block;" class="clone" aria-hidden="true">
                <div class="sliderwrapper">
                    <a href="http://google.ca">
                        <img class="slideShowImages" src="/sites/MCISPOTestSite/SiteAssets/Branding/images/LandingPageSlider/CarImage1.png" alt="missing" draggable="false">
                    </a>
                    <div class="container">
                        <div class="content">
                            <h2>TEst</h2>
                            <div>
                                <p></p>
                            </div>
                            <p><a href="http://google.ca" class="use-sprite">
                                <!--<pointfire:control sectionid="Navigationfrench" language="1036" method="hide" />--><span style="display:none;"><span style="display:none;">En savoir plus</span></span><!--<pointfire:control endsection="Navigationfrench" />-->
                           		<!--<pointfire:control sectionid="Navigationenglish" language="1033" method="hide" />-->Find out more<!--<pointfire:control endsection="Navigationenglish" />--> 
                            <span class="sprite-link-arrow">  </span></a></p>
                        </div>
                    </div>
                </div>
            </li></ul></div><ol class="flex-control-nav flex-control-paging"><li><a class="">1</a></li><li><a class="">2</a></li><li><a class="flex-active">3</a></li></ol><ul class="flex-direction-nav"><li><a class="flex-prev" href="#">Previous</a></li><li><a class="flex-next" href="#">Next</a></li></ul></div>
            <!--For tiles-->
            <div id="blankRow1"></div>
            <div id="second-row-home" class="">
				<div id="scriptWPQ6">
					<div id="WPQ6-GettingStarted" class="ms-promlink-root">
						<div id="promotedlinksheader_WPQ6" class="ms-promlink-header">
							<span><h2 class="ms-promlink-parttitle ms-webpart-titleText"></h2></span>
							<span><a id="promotedlinks_hide_WPQ6" href="#" title="Remove these tiles from the page and access them later from the Site menu." class="ms-commandLink ms-blog-command">Remove this</a></span>
							<span id="promotedlinkspaging_WPQ6" class="ms-promlink-headerNav" style="display: none;"><a class="ms-commandLink ms-promlink-button ms-promlink-button-disabled" id="promotedlinks_prev_WPQ6" title="Previous" href="#"><span class="ms-promlink-button-image"><img class="ms-promlink-button-left-disabled" id="promotedlinks_prev_WPQ6img" alt="Previous" src="/sites/MCISPOTestSite/SiteAssets/Branding/images/spcommon.png?rev=40"></span></a><span class="ms-promlink-button-inner" id="promotedlinks_content_WPQ6"></span><a class="ms-commandLink ms-promlink-button ms-promlink-button-disabled" id="promotedlinks_next_WPQ6" title="Next" href="#"><span class="ms-promlink-button-image"><img class="ms-promlink-button-right-disabled" id="promotedlinks_next_WPQ6img" alt="Next" src="/sites/MCISPOTestSite/SiteAssets/Branding/images/spcommon.png?rev=40"></span></a></span></div><table style="width:100%; table-layout:fixed;"><tbody><tr><td><div id="promotedlinksbody_WPQ6" class="ms-promlink-body"><div id="Tile_WPQ6_6_1" style="width: 160px; height: 160px;" class="ms-tileview-tile-root"><div id="Tile_WPQ6_6_2" style="width: 150px; height: 150px;" aria-haspopup="true" class="ms-tileview-tile-content"><a target="_blank" id="Tile_WPQ6_6_3" href="http://www.mazda247.ca/" target="_blank" onclick="PreventDefaultNavigation(); return false;" hrefaction="http://www.mazda247.ca/" clickaction="null"><div style="height:100%"><span style="height:150px; width:150px; position:relative; display:inline-block; overflow:hidden;"><img src="https://mazdausa.sharepoint.com/sites/MCISPOTestSite/SiteAssets/Branding/images/Tiles/Accessories.png" style="left:-0px; top:-0px; position:absolute;" onerror="return SP.UI.PromotedLinks.BackgroundImage.OnLoadError(this);" alt="Data Asset Manager"></span><div id="Tile_WPQ6_6_4" style="width: 150px; height: 150px; top: 100px; left: 0px;" offy="100" class="ms-tileview-tile-detailsBox"><ul class="ms-noList ms-tileview-tile-detailsListMedium"><li id="Tile_WPQ6_6_5" collapsed="ms-tileview-tile-titleMedium ms-tileview-tile-titleMediumCollapsed" expanded="ms-tileview-tile-titleMedium ms-tileview-tile-titleMediumExpanded" class="ms-tileview-tile-titleMedium ms-tileview-tile-titleMediumCollapsed"><div collapsed="ms-tileview-tile-titleTextMediumCollapsed" expanded="ms-tileview-tile-titleTextMediumExpanded" class="ms-tileview-tile-titleTextMediumCollapsed">Data Asset Manager</div></li><li title="Data Asset Manager" id="Tile_WPQ6_6_6" class="ms-tileview-tile-descriptionMedium"></li></ul></div></div></a></div></div><div id="Tile_WPQ6_5_1" style="width: 160px; height: 160px;" class="ms-tileview-tile-root"><div id="Tile_WPQ6_5_2" style="width: 150px; height: 150px;" aria-haspopup="true" class="ms-tileview-tile-content"><a target="_blank" id="Tile_WPQ6_5_3" href="#" onclick="PreventDefaultNavigation(); return false;" hrefaction="https://one.mazda.ca/Pages/PerformanceHubInterMdt.aspx" clickaction="null"><div style="height:100%"><span style="height:150px; width:150px; position:relative; display:inline-block; overflow:hidden;"><img src="https://mazdausa.sharepoint.com/sites/MCISPOTestSite/SiteAssets/Branding/images/Tiles/Accessories.png" style="left:-0px; top:-0px; position:absolute;" onerror="return SP.UI.PromotedLinks.BackgroundImage.OnLoadError(this);" alt="Performance Hub"></span><div id="Tile_WPQ6_5_4" style="width: 150px; height: 150px; top: 100px; left: 0px;" offy="100" class="ms-tileview-tile-detailsBox"><ul class="ms-noList ms-tileview-tile-detailsListMedium"><li id="Tile_WPQ6_5_5" collapsed="ms-tileview-tile-titleMedium ms-tileview-tile-titleMediumCollapsed" expanded="ms-tileview-tile-titleMedium ms-tileview-tile-titleMediumExpanded" class="ms-tileview-tile-titleMedium ms-tileview-tile-titleMediumCollapsed"><div collapsed="ms-tileview-tile-titleTextMediumCollapsed" expanded="ms-tileview-tile-titleTextMediumExpanded" class="ms-tileview-tile-titleTextMediumCollapsed">Performance Hub</div></li><li title="Performance Hub" id="Tile_WPQ6_5_6" class="ms-tileview-tile-descriptionMedium"></li></ul></div></div></a></div></div><div id="Tile_WPQ6_4_1" style="width: 160px; height: 160px;" class="ms-tileview-tile-root"><div id="Tile_WPQ6_4_2" style="width: 150px; height: 150px;" aria-haspopup="true" class="ms-tileview-tile-content"><a target="_blank" id="Tile_WPQ6_4_3" href="#" hrefaction="https://learningmanager.adobe.com/splogin?accountId=112952&amp;isExternal=false" clickaction="null" onclick="PreventDefaultNavigation(); return false;"><div style="height:100%"><span style="height:150px; width:150px; position:relative; display:inline-block; overflow:hidden;"><img src="https://mazdausa.sharepoint.com/sites/MCISPOTestSite/SiteAssets/Branding/images/Tiles/Accessories.png" style="left:-0px; top:-0px; position:absolute;" onerror="return SP.UI.PromotedLinks.BackgroundImage.OnLoadError(this);" alt="Mazda Brand Academy"></span><div id="Tile_WPQ6_4_4" style="width: 150px; height: 150px; top: 100px; left: 0px;" offy="100" class="ms-tileview-tile-detailsBox"><ul class="ms-noList ms-tileview-tile-detailsListMedium"><li id="Tile_WPQ6_4_5" collapsed="ms-tileview-tile-titleMedium ms-tileview-tile-titleMediumCollapsed" expanded="ms-tileview-tile-titleMedium ms-tileview-tile-titleMediumExpanded" class="ms-tileview-tile-titleMedium ms-tileview-tile-titleMediumCollapsed"><div collapsed="ms-tileview-tile-titleTextMediumCollapsed" expanded="ms-tileview-tile-titleTextMediumExpanded" class="ms-tileview-tile-titleTextMediumCollapsed">Mazda Brand Academy</div></li><li title="Mazda Brand Academy" id="Tile_WPQ6_4_6" class="ms-tileview-tile-descriptionMedium"></li></ul></div></div></a></div></div><div id="Tile_WPQ6_3_1" style="width: 160px; height: 160px;" class="ms-tileview-tile-root"><div id="Tile_WPQ6_3_2" style="width: 150px; height: 150px;" aria-haspopup="true" class="ms-tileview-tile-content"><a target="_blank" id="Tile_WPQ6_3_3" href="#" hrefaction="https://one.mazda.ca/Service/Accessories" clickaction="null" onclick="PreventDefaultNavigation(); return false;"><div style="height:100%"><span style="height:150px; width:150px; position:relative; display:inline-block; overflow:hidden;"><img src="https://mazdausa.sharepoint.com/sites/MCISPOTestSite/SiteAssets/Branding/images/Tiles/Accessories.png" style="left:-0px; top:-0px; position:absolute;" onerror="return SP.UI.PromotedLinks.BackgroundImage.OnLoadError(this);" alt="Accessories"></span><div id="Tile_WPQ6_3_4" style="width: 150px; height: 150px; top: 100px; left: 0px;" offy="100" class="ms-tileview-tile-detailsBox"><ul class="ms-noList ms-tileview-tile-detailsListMedium"><li id="Tile_WPQ6_3_5" collapsed="ms-tileview-tile-titleMedium ms-tileview-tile-titleMediumCollapsed" expanded="ms-tileview-tile-titleMedium ms-tileview-tile-titleMediumExpanded" class="ms-tileview-tile-titleMedium ms-tileview-tile-titleMediumCollapsed"><div collapsed="ms-tileview-tile-titleTextMediumCollapsed" expanded="ms-tileview-tile-titleTextMediumExpanded" class="ms-tileview-tile-titleTextMediumCollapsed">Accessories</div></li><li title="Accessories" id="Tile_WPQ6_3_6" class="ms-tileview-tile-descriptionMedium"></li></ul></div></div></a></div></div><div id="Tile_WPQ6_2_1" style="width: 160px; height: 160px;" class="ms-tileview-tile-root"><div id="Tile_WPQ6_2_2" style="width: 150px; height: 150px;" aria-haspopup="true" class="ms-tileview-tile-content"><a target="_blank" id="Tile_WPQ6_2_3" href="#" onclick="PreventDefaultNavigation(); return false;" hrefaction="https://one.mazda.ca/Pages/FrequentlyUsedDealerForms.aspx" clickaction="null"><div style="height:100%"><span style="height:150px; width:150px; position:relative; display:inline-block; overflow:hidden;"><img src="https://mazdausa.sharepoint.com/sites/MCISPOTestSite/SiteAssets/Branding/images/Tiles/Accessories.png" style="left:-0px; top:-0px; position:absolute;" onerror="return SP.UI.PromotedLinks.BackgroundImage.OnLoadError(this);" alt="Forms"></span><div id="Tile_WPQ6_2_4" style="width: 150px; height: 150px; top: 100px; left: 0px;" offy="100" class="ms-tileview-tile-detailsBox"><ul class="ms-noList ms-tileview-tile-detailsListMedium"><li id="Tile_WPQ6_2_5" collapsed="ms-tileview-tile-titleMedium ms-tileview-tile-titleMediumCollapsed" expanded="ms-tileview-tile-titleMedium ms-tileview-tile-titleMediumExpanded" class="ms-tileview-tile-titleMedium ms-tileview-tile-titleMediumCollapsed"><div collapsed="ms-tileview-tile-titleTextMediumCollapsed" expanded="ms-tileview-tile-titleTextMediumExpanded" class="ms-tileview-tile-titleTextMediumCollapsed">Forms</div></li><li title="Forms" id="Tile_WPQ6_2_6" class="ms-tileview-tile-descriptionMedium"></li></ul></div></div></a></div></div><div id="Tile_WPQ6_1_1" style="width: 160px; height: 160px;" class="ms-tileview-tile-root"><div id="Tile_WPQ6_1_2" style="width: 150px; height: 150px;" aria-haspopup="true" class="ms-tileview-tile-content"><a target="_blank" id="Tile_WPQ6_1_3" href="#" onclick="PreventDefaultNavigation(); return false;" hrefaction="https://one.mazda.ca/Pages/businessManagement.aspx" clickaction="null"><div style="height:100%"><span style="height:150px; width:150px; position:relative; display:inline-block; overflow:hidden;"><img src="https://mazdausa.sharepoint.com/sites/MCISPOTestSite/SiteAssets/Branding/images/Tiles/Accessories.png" style="left:-0px; top:-0px; position:absolute;" onerror="return SP.UI.PromotedLinks.BackgroundImage.OnLoadError(this);" alt="Business Management"></span><div id="Tile_WPQ6_1_4" style="width: 150px; height: 150px; top: 100px; left: 0px;" offy="100" class="ms-tileview-tile-detailsBox"><ul class="ms-noList ms-tileview-tile-detailsListMedium"><li id="Tile_WPQ6_1_5" collapsed="ms-tileview-tile-titleMedium ms-tileview-tile-titleMediumCollapsed" expanded="ms-tileview-tile-titleMedium ms-tileview-tile-titleMediumExpanded" class="ms-tileview-tile-titleMedium ms-tileview-tile-titleMediumCollapsed"><div collapsed="ms-tileview-tile-titleTextMediumCollapsed" expanded="ms-tileview-tile-titleTextMediumExpanded" class="ms-tileview-tile-titleTextMediumCollapsed">Business Management</div></li><li title="Business Management" id="Tile_WPQ6_1_6" class="ms-tileview-tile-descriptionMedium"></li></ul></div></div></a></div></div></div></td></tr></tbody></table></div></div>
							</div>
            `;
  }

  protected onInit(): Promise<void> {
    SPComponentLoader.loadCss('https://mazdausa.sharepoint.com/sites/MCISPOTestSite/SiteAssets/NameTag/css/bootstrap.min.css');
    SPComponentLoader.loadCss('https://mazdausa.sharepoint.com/sites/MCISPOTestSite/SiteAssets/Branding/css/styles/Themable/corev15.css');
    SPComponentLoader.loadCss('https://mazdausa.sharepoint.com/sites/MCISPOTestSite/SiteAssets/Branding/css/OneMazda_PubCollab_Home.css');
    SPComponentLoader.loadCss('https://mazdausa.sharepoint.com/sites/MCISPOTestSite/SiteAssets/Branding/css/NewsSlider.css');
    SPComponentLoader.loadCss('https://mazdausa.sharepoint.com/sites/MCISPOTestSite/SiteAssets/Branding/css/corev15.css');
    SPComponentLoader.loadScript("https://mazdausa.sharepoint.com/sites/MCISPOTestSite/SiteAssets/Branding/js/jquery-1.10.2.min.js",{globalExportsName:'jquery'}).then(()=>{
    SPComponentLoader.loadScript("https://mazdausa.sharepoint.com/sites/MCISPOTestSite/SiteAssets/Branding/js/slider/jquery.flexslider.js",{globalExportsName:'sliderJs'}).then(()=>{
    SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCISPOTestSite/SiteAssets/Branding/js/slider/NewsSlider.js');  
    });
  });
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
