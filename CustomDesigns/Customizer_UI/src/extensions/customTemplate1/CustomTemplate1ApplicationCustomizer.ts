import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';
//import { Dialog } from '@microsoft/sp-dialog';
//import { SPComponentLoader } from '@microsoft/sp-loader';

import * as strings from 'CustomTemplate1ApplicationCustomizerStrings';
import styles from './AppCustomizer.module.scss';
//import { escape } from '@microsoft/sp-lodash-subset';
import { override } from '@microsoft/decorators';


const LOG_SOURCE: string = 'CustomTemplate1ApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ICustomTemplate1ApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: "Test Message 101";
  cssurl: string;
  Top: string;
  Bottom: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class CustomTemplate1ApplicationCustomizer
  extends BaseApplicationCustomizer<ICustomTemplate1ApplicationCustomizerProperties> {
     // These have been added
  private _topPlaceholder: PlaceholderContent | undefined;
  //private _bottomPlaceholder: PlaceholderContent | undefined;

@override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
    //SPComponentLoader.loadCss('https://1djbv0.sharepoint.com/sites/Devansh1/SiteAssets/CSS/DemoCss1.css');
    //SPComponentLoader.loadScript("https://mazdausa.sharepoint.com/sites//SiteAssets/Branding/js/jquery.js",{globalExportsName:'jqueryCustomizer'}).then(()=>{
    //SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCISPOTest/SiteAssets/Branding/js/branding.js');
  //});
const cssUrl: string = this.properties.cssurl;
if(cssUrl)
{
  const head: any = document.getElementsByTagName("head")[0] || document.documentElement;
  const customStyle: HTMLLinkElement = document.createElement("link");
  customStyle.href=cssUrl;
  customStyle.rel="stylesheet";
  customStyle.type="text/css";
  head.insertAdjacentElement("beforeEnd", customStyle);
}
    let message: string = this.properties.testMessage;
    if (!message) {
      message = '(No properties were provided.)';
    }

    //Dialog.alert(`Hello from ${strings.Title}:\n\n${message}`).catch(() => {
      /* handle error */
    //});
    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);

    return Promise.resolve();
  }
  private _renderPlaceHolders(): void {
    console.log("HelloWorldApplicationCustomizer._renderPlaceHolders()");
    console.log(
      "Available placeholders: ",
      this.context.placeholderProvider.placeholderNames
        .map(name => PlaceholderName[name])
        .join(", ")
    );
  
    // Handling the top placeholder
    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        { onDispose: this._onDispose }
      );
  
      // The extension should not assume that the expected placeholder is available.
      if (!this._topPlaceholder) {
        console.error("The expected placeholder (Top) was not found.");
        return;
      }
  
      if (this.properties) {
        let topString: string = this.properties.Top;
        if (!topString) {
          topString = "MZD.Com";
        }
  
        if (this._topPlaceholder.domElement) {
          this._topPlaceholder.domElement.innerHTML = `
          <div class="${styles.app}">
            <div class="${styles.top}">
              <i class="ms-Icon ms-Icon--Info" aria-hidden="true"></i> <img id="topIcon" src="/sites/MCISPOTest/SiteAssets/Branding/images/Mazda_Logo_Horizontal_CMYK_250.png">
            </div>
            <div id="homePageLink">
            <!--<a href="/"><img id="mainPageLogo" src="/sites/MCISPOTest/SiteAssets/Branding/images/onemazda2.png"></a>-->
            <a id="anchorSiteName" href="/sites/MCISPOTest">MCISPOTest</a>
            </div>
            <div id="topMenu">
            <ul id="topMenuItems">
            <li><a href="#" data-interception="off">eBulletins</a></li> 
            <li><a href="#" data-interception="off">Dealer Affairs</a></li>
            <li><a href="#" data-interception="off">Administration</a></li> 
            <li><a href="#" data-interception="off">Regions</a></li>
            <li><a href="#" data-interception="off">Sales</a></li> 
            <li><a href="#" data-interception="off">Marketing</a></li>
            <li><a href="#" data-interception="off">Parts</a></li> 
            <li><a href="#" data-interception="off">Service</a></li>
            <li><a href="#" data-interception="off">Customer Insight</a></li>                     
          </ul>
            </div>            
          </div>`;
        }
      }
    }
  
    // Handling the bottom placeholder
    /*if (!this._bottomPlaceholder) {
      this._bottomPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Bottom,
        { onDispose: this._onDispose }
      );
  
      // The extension should not assume that the expected placeholder is available.
      if (!this._bottomPlaceholder) {
        console.error("The expected placeholder (Bottom) was not found.");
        return;
      }
  
      if (this.properties) {
        let bottomString: string = this.properties.Bottom;
        if (!bottomString) {
          bottomString = "www.mzd.com";
        }
  
        if (this._bottomPlaceholder.domElement) {
          this._bottomPlaceholder.domElement.innerHTML = `
          <div class="${styles.app}">
            <div class="${styles.bottom}">
              <i class="ms-Icon ms-Icon--Info" aria-hidden="true"></i> ${escape(
                bottomString
              )}
            </div>
          </div>`;
        }
      }
    }*/
  }
  private _onDispose(): void {
    console.log('[HelloWorldApplicationCustomizer._onDispose] Disposed custom top and bottom placeholders.');
  }
}
