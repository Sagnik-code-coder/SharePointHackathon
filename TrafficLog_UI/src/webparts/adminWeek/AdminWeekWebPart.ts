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

import * as strings from 'AdminWeekWebPartStrings';
import AdminWeek from './components/AdminWeek';
import { IAdminWeekProps } from './components/IAdminWeekProps';

export interface IAdminWeekWebPartProps {
  description: string;
}

export default class AdminWeekWebPart extends BaseClientSideWebPart<IAdminWeekWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IAdminWeekProps> = React.createElement(
      AdminWeek,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
    this.domElement.innerHTML =
    `
    <div id="loader" class="centerloader"></div>
    <div class="body-content">
        <div class="container-fluid">
            <div style="text-align:right">
                <a target="_blank" href="https://mazdausa.sharepoint.com/sites/MCIOneMazdaDev/Trafficlog/TrafficLog%20Documents/Traffic%20Calender" tkey="TrafficPDFUpload"
                    class="ui-tabs-anchor">Upload Traffic Calendar (PDF)</a>
            </div>
            <div>
                <div class="alert alert-info" tkey="UserInaccessibility" id="UserInaccessibilityAlert"
                    style="display:none">The current user is not allowed to visit this page.</div>
                <div id="grids">
                    <div id="dealersubmit">
                    <span style="display: none;" id="Currculture">`+ this.context.pageContext.cultureInfo.currentCultureName +`</span>
                        <h1 class="heading" tkey="TrafficLogSystemAdministration">Traffic Log &amp; System
                            Administration</h1>
                        <div class="alert alert-info" tkey="DealerSubmitMessage" id="DealerSubmitMessage"
                            style="display:none">Your request has been successfully submitted.</div>
                        <div class="alert alert-info" tkey="DealerFailureMessage" id="DealerFailureMessage"
                            style="display:none">Your Submission is not successfully submitted.Please contact support
                            team.</div>

                        <div wfd-id="9" class="filter_area">


                            <div style="width:100%; margin:0 auto;">
                                <table id="reportingWeekCalender"
                                    class="table table-striped table-bordered dt-responsive nowrap data-entry"
                                    cellspacing="0">
                                    <thead>
                                        <tr>
                                            <th> Week ID</th>
                                            <th> Week Number</th>
                                            <th> Week Description EN</th>
                                            <th> Week Description FR</th>
                                            <th> Week Start Date</th>
                                            <th> Week End Date</th>
                                            <th> Month Number</th>
                                            <th> Year Number</th>
                                            <th> Submission Start Date</th>
                                            <th> Submission End Date</th>
                                            <th> Submission First DeadLine</th>
                                            <th> Submission Description EN</th>
                                            <th> Submission Description FR</th>
                                        </tr>
                                    </thead>
                                    <tbody>

                                    </tbody>
                                </table>
                            </div>
                            <p style="text-align:right">
                                <button id="dealersubmitbtn" tkey="Submit" type="submit" value="Submit"
                                    class="btn btn-dark primary_btn">Submit</button>
                                <button id="reset" tkey="Reset"
                                    class="btn btn-outline-secondary secondary_btn">Reset</button>
                            </p>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>






    <script src="bootstrap.js"></script>
    <script src="respond.js"></script>
    <script src="Scripts.js"></script>
    <script src="notify.min.js"></script>
    <script src="jquery-1.10.2.min.js"></script>

    <script src="jquery.validate.min.js"></script>
    <script src="jquery.validate.unobtrusive.min.js"></script>

    <script src="https://cdnjs.cloudflare.com/ajax/libs/jsgrid/1.5.3/jsgrid.min.js"></script>
    <script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>



    <script src="reportcalender.js"></script>
    <style type="text/css">
        .ms-webpart-titleText {
            display: none;
        }
    </style>


    
    `
  }

  protected onInit(): Promise<void> {
    SPComponentLoader.loadCss('https://mazdausa.sharepoint.com/sites/MCIOneMazdaDev/trafficlog/SiteAssets/css/jsgrid.min.css');
    // SPComponentLoader.loadCss('https://mazdausa.sharepoint.com/sites/MCIOneMazdaStg/trafficlog/SiteAssets/css/jsgrid.min.css'); For Staging, uncomment the below line!
    // SPComponentLoader.loadCss('https://mazdausa.sharepoint.com/sites/MCIOneMazda/trafficlog/SiteAssets/css/jsgrid.min.css'); For Production, uncomment the below line!
        SPComponentLoader.loadCss('https://mazdausa.sharepoint.com/sites/MCIOneMazdaDev/trafficlog/SiteAssets/css/jsgrid-theme.min.css');
        // SPComponentLoader.loadCss('https://mazdausa.sharepoint.com/sites/MCIOneMazdaStg/trafficlog/SiteAssets/css/jsgrid-theme.min.css'); For Staging, uncomment the below line!
        // SPComponentLoader.loadCss('https://mazdausa.sharepoint.com/sites/MCIOneMazda/trafficlog/SiteAssets/css/jsgrid-theme.min.css'); For Production, uncomment the below line!
        SPComponentLoader.loadCss('https://mazdausa.sharepoint.com/sites/MCIOneMazdaDev/trafficlog/SiteAssets/css/bootstrap.min.css');
        // SPComponentLoader.loadCss('https://mazdausa.sharepoint.com/sites/MCIOneMazdaStg/trafficlog/SiteAssets/css/bootstrap.min.css'); For Staging, uncomment the below line!
        // SPComponentLoader.loadCss('https://mazdausa.sharepoint.com/sites/MCIOneMazda/trafficlog/SiteAssets/css/bootstrap.min.css'); For Production, uncomment the below line!
        SPComponentLoader.loadCss('https://mazdausa.sharepoint.com/sites/MCIOneMazdaDev/trafficlog/SiteAssets/css/ui.css');
        // SPComponentLoader.loadCss('https://mazdausa.sharepoint.com/sites/MCIOneMazdaStg/trafficlog/SiteAssets/css/ui.css'); For Staging, uncomment the below line!
        // SPComponentLoader.loadCss('https://mazdausa.sharepoint.com/sites/MCIOneMazda/trafficlog/SiteAssets/css/ui.css'); For Production, uncomment the below line!
    SPComponentLoader.loadCss('https://mazdausa.sharepoint.com/sites/MCIOneMazdaDev/trafficlog/SiteAssets/css/bootstrap.css');
    // SPComponentLoader.loadCss('https://mazdausa.sharepoint.com/sites/MCIOneMazdaStg/trafficlog/SiteAssets/css/bootstrap.css'); For Staging, uncomment the below line!
    // SPComponentLoader.loadCss('https://mazdausa.sharepoint.com/sites/MCIOneMazda/trafficlog/SiteAssets/css/bootstrap.css'); For Production, uncomment the below line!
    SPComponentLoader.loadCss('https://mazdausa.sharepoint.com/sites/MCIOneMazdaDev/trafficlog/SiteAssets/css/Site.css');
    // SPComponentLoader.loadCss('https://mazdausa.sharepoint.com/sites/MCIOneMazdaStg/trafficlog/SiteAssets/css/Site.css'); For Staging, uncomment the below line!
    // SPComponentLoader.loadCss('https://mazdausa.sharepoint.com/sites/MCIOneMazda/trafficlog/SiteAssets/css/Site.css'); For Production, uncomment the below line!
    SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/bootstrap-datetimepicker/4.17.44/css/bootstrap-datetimepicker.min.css'); //--
    SPComponentLoader.loadCss('https://mazdausa.sharepoint.com/sites/MCIOneMazdaDev/trafficlog/SiteAssets/css/jquery.datetimepicker.min.css');
    // SPComponentLoader.loadCss('https://mazdausa.sharepoint.com/sites/MCIOneMazdaStg/trafficlog/SiteAssets/css/jquery.datetimepicker.min.css'); For Staging, uncomment the below line!
    // SPComponentLoader.loadCss('https://mazdausa.sharepoint.com/sites/MCIOneMazda/trafficlog/SiteAssets/css/jquery.datetimepicker.min.css'); For Production, uncomment the below line!
    
    SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCIOneMazdaDev/trafficlog/SiteAssets/js/jquery-3.6.0.min.js')
    // SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCIOneMazdaStg/trafficlog/SiteAssets/js/jquery-3.6.0.min.js') For Staging, uncomment the below line!
    // SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCIOneMazda/trafficlog/SiteAssets/js/jquery-3.6.0.min.js') For Production, uncomment the below line!
    .then(() => {
        return SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCIOneMazdaDev/trafficlog/SiteAssets/js/jquery-ui.min.js');
        // return SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCIOneMazdaStg/trafficlog/SiteAssets/js/jquery-ui.min.js'); For Staging, uncomment the below line!
        // return SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCIOneMazda/trafficlog/SiteAssets/js/jquery-ui.min.js'); For Production, uncomment the below line!

    })
    .then(() =>{
        return SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCIOneMazdaDev/trafficlog/SiteAssets/js/jquery.datetimepicker.js');
        // return SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCIOneMazdaStg/trafficlog/SiteAssets/js/jquery.datetimepicker.js'); For Staging, uncomment the below line!
        // return SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCIOneMazda/trafficlog/SiteAssets/js/jquery.datetimepicker.js'); For Production, uncomment the below line!
    })  
    .then(() => {
        return SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCIOneMazdaDev/trafficlog/SiteAssets/js/jsgrid.min.js');
        // return SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCIOneMazdaStg/trafficlog/SiteAssets/js/jsgrid.min.js'); For Staging, uncomment the below line!
        // return SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCIOneMazda/trafficlog/SiteAssets/js/jsgrid.min.js'); For Production, uncomment the below line!
    })
    .then(()=>{
      return SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCIOneMazdaDev/trafficlog/SiteAssets/js/lang.js');
  })
    .then(() => {
        return SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCIOneMazdaDev/trafficlog/SiteAssets/js/moment.js');
        // return SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCIOneMazdaStg/trafficlog/SiteAssets/js/moment.js'); For Staging, uncomment the below line!
        // return SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCIOneMazda/trafficlog/SiteAssets/js/moment.js'); For Production, uncomment the below line!
    })
    .then(() => {
        //return SPComponentLoader.loadScript('https://y3mbk.sharepoint.com/sites/SharePointCRUD/SiteAssets/TrafficLog/js/jquery.datetimepicker.js');
    })
    .then(() => {
        return SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCIOneMazdaDev/trafficlog/SiteAssets/js/ReportCalender1.js');
        // return SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCIOneMazdaStg/trafficlog/SiteAssets/js/ReportCalender1.js'); For Staging, uncomment the below line!
        // return SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCIOneMazda/trafficlog/SiteAssets/js/ReportCalender1.js'); For Production, uncomment the below line!
    })
    .catch((error) => {
        console.error('Error loading scripts:', error);
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
