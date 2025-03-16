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
import * as jQuery from 'jquery';
import 'jqueryui';
 
 
 
import * as strings from 'DealerEntryWebPartStrings';
import DealerEntry from './components/DealerEntry';
import { IDealerEntryProps } from './components/IDealerEntryProps';
import { SPHttpClient, SPHttpClientResponse} from '@microsoft/sp-http';
 
export interface IDealerEntryWebPartProps {
    description: string;
}
// eslint-disable-next-line @typescript-eslint/no-unused-vars
 
 
 
export default class DealerEntryWebPart extends BaseClientSideWebPart<IDealerEntryWebPartProps> {
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
        SPComponentLoader.loadCss('https://mazdausa.sharepoint.com/sites/MCIOneMazdaDev/trafficlog/SiteAssets/css/responsive.bootstrap.min.css');
        // SPComponentLoader.loadCss('https://mazdausa.sharepoint.com/sites/MCIOneMazdaStg/trafficlog/SiteAssets/css/responsive.bootstrap.min.css'); For Staging, uncomment the below line!
        // SPComponentLoader.loadCss('https://mazdausa.sharepoint.com/sites/MCIOneMazda/trafficlog/SiteAssets/css/responsive.bootstrap.min.css'); For Production, uncomment the below line!
        SPComponentLoader.loadCss('https://mazdausa.sharepoint.com/sites/MCIOneMazdaDev/trafficlog/SiteAssets/css/bootstrap.css');
        // SPComponentLoader.loadCss('https://mazdausa.sharepoint.com/sites/MCIOneMazdaStg/trafficlog/SiteAssets/css/bootstrap.css'); For Staging, uncomment the below line!
        // SPComponentLoader.loadCss('https://mazdausa.sharepoint.com/sites/MCIOneMazda/trafficlog/SiteAssets/css/bootstrap.css'); For Production, uncomment the below line!
        SPComponentLoader.loadCss('https://mazdausa.sharepoint.com/sites/MCIOneMazdaDev/trafficlog/SiteAssets/css/Site.css');
        // SPComponentLoader.loadCss('https://mazdausa.sharepoint.com/sites/MCIOneMazdaStg/trafficlog/SiteAssets/css/Site.css'); For Staging, uncomment the below line!
        // SPComponentLoader.loadCss('https://mazdausa.sharepoint.com/sites/MCIOneMazda/trafficlog/SiteAssets/css/Site.css'); For Production, uncomment the below line!
       
        SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCIOneMazdaDev/trafficlog/SiteAssets/js/jquery-1.10.2.js')
        // SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCIOneMazdaStg/trafficlog/SiteAssets/js/jquery-1.10.2.js') For Staging, uncomment the below line!
        // SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCIOneMazda/trafficlog/SiteAssets/js/jquery-1.10.2.js') For Production, uncomment the below line!
        .then(() => {
            return SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCIOneMazdaDev/trafficlog/SiteAssets/js/jsgrid.min.js');
            // return SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCIOneMazdaStg/trafficlog/SiteAssets/js/jsgrid.min.js'); For Staging, uncomment the below line!
            // return SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCIOneMazda/trafficlog/SiteAssets/js/jsgrid.min.js'); For Production, uncomment the below line!
        })
        .then(()=>{
            return SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCIOneMazdaDev/trafficlog/SiteAssets/js/lang.js');
        })
        .then(() => {
            return SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCIOneMazdaDev/trafficlog/SiteAssets/js/jquery-ui.js');
            // return SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCIOneMazdaStg/trafficlog/SiteAssets/js/jquery-ui.js'); For Staging, uncomment the below line!
            // return SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCIOneMazda/trafficlog/SiteAssets/js/jquery-ui.js'); For Production, uncomment the below line!
        })
        .then(()=>{
            //return SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/_layouts/15/sp.runtime.js');
        })
        .then(()=>{
            //return SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/_layouts/15/sp.js');
        })
        .then(()=>{
            return SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCIOneMazdaDev/trafficlog/SiteAssets/js/lang.js');
        })
        .then(() => {
            return SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCIOneMazdaDev/trafficlog/SiteAssets/js/dealerentry.js');
            // return SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCIOneMazdaStg/trafficlog/SiteAssets/js/dealerentry.js'); For Staging, uncomment the below line!
            // return SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCIOneMazda/trafficlog/SiteAssets/js/dealerentry.js'); For Production, uncomment the below line!
        })
        .then(() => {
            return SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCIOneMazdaDev/trafficlog/SiteAssets/js/dealerdetailreport.js');
            // return SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCIOneMazdaStg/trafficlog/SiteAssets/js/dealerdetailreport.js'); For Staging, uncomment the below line!
            // return SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCIOneMazda/trafficlog/SiteAssets/js/dealerdetailreport.js'); For Production, uncomment the below line!
        })
        .then(() => {
            return SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCIOneMazdaDev/trafficlog/SiteAssets/js/csvExport.js');
            // return SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCIOneMazdaStg/trafficlog/SiteAssets/js/csvExport.js'); For Staging, uncomment the below line!
            // return SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCIOneMazda/trafficlog/SiteAssets/js/csvExport.js'); For Production, uncomment the below line!
        })
       
 
        return this._getEnvironmentMessage().then(message => {
            this._environmentMessage = message;
        });
    }
    private _isDarkTheme: boolean = false;
    private _environmentMessage: string = '';
 
 
 
    /*
    public componentDidMount(): void {
        jQuery("#btnMTDTrend").on('click', () => {
            jQuery("#MTDTrend").toggle(); // Toggle the visibility of the table
        });
    }
    */
 
 
 
    public render(): void {
        const element: React.ReactElement<IDealerEntryProps> = React.createElement(
            DealerEntry,
            {
                description: this.properties.description,
                isDarkTheme: this._isDarkTheme,
                environmentMessage: this._environmentMessage,
                hasTeamsContext: !!this.context.sdks.microsoftTeams,
                userDisplayName: this.context.pageContext.user.displayName
            }
        );
 
        ReactDom.render(element, this.domElement);
        //this.domElement.innerHTML =
        const tabsHTML: string =
            `<div id="loader" class="centerloader"></div>
    <div class="body-content">
       
        <div class="container-fluid">
            <div style="text-align:right" id="divDownload">
                <!--<a href="/trafficlog/_layouts/download.aspx?SourceUrl=/trafficlog/Documents/Traffic%20Calendar.pdf" download="Traffic Calendar.pdf" tkey="TrafficPDFDownload" class="ui-tabs-anchor">Download Traffic Calendar (PDF)</a>-->
                <!--<a target="_blank" href="/trafficlog/Documents/Traffic%20Calendar.pdf" tkey="TrafficPDFDownload" class="ui-tabs-anchor">Download Traffic Calendar (PDF)</a>-->
                <!--<a id="downloadPDF" href="#" tkey="TrafficPDFDownload" class="ui-tabs-anchor">Download Traffic Calendar (PDF)</a>-->
            </div>
            <div id="tabs" class="ui-tabs ui-corner-all ui-widget ui-widget-content">
            <span style="display: none;" id="Currculture">`+ this.context.pageContext.cultureInfo.currentCultureName +`</span>
                <ul role="tablist" class="ui-tabs-nav ui-corner-all ui-helper-reset ui-helper-clearfix ui-widget-header">
                    <li role="tab" tabindex="0" class="ui-tabs-tab ui-corner-top ui-state-default ui-tab ui-tabs-active ui-state-active" aria-controls="grids" aria-labelledby="ui-id-1" aria-selected="true" aria-expanded="true"><a href="#grids" tkey="DealerEntry" role="presentation" tabindex="-1" class="ui-tabs-anchor" id="ui-id-1">Clubhouse Entry</a></li>
                    <li role="tab" tabindex="-1" class="ui-tabs-tab ui-corner-top ui-state-default ui-tab" aria-controls="Summaries" aria-labelledby="ui-id-2" aria-selected="false" aria-expanded="false"><a href="#Summaries" tkey="DealersInformationReport" role="presentation" tabindex="-1" class="ui-tabs-anchor" id="ui-id-2">Clubhouse Information Report</a></li>
                </ul>
                <div id="grids" aria-labelledby="ui-id-1" role="tabpanel" class="ui-tabs-panel ui-corner-bottom ui-widget-content" aria-hidden="false" style="display: block;">
                    <div id="dealersubmit">
                        <h1 class="heading" tkey="TrafficWritesForecastEntry">Traffic, Writes &amp; Forecast Entry</h1>
                        <div class="alert alert-info" tkey="DealerMessage" id="DealerMessage" style="display:none">Your Submission will be after the first submission deadline.</div>
                        <div class="alert alert-info" tkey="OutsideTheSubmissionWindow" id="OutsideTheSubmissionWindow" style="display:none">Sorry, After or Before the submission dates entry is not allowed.</div>
                        <div class="alert alert-info" tkey="DealerSubmitMessage" id="DealerSubmitMessage" style="display:none">Your request has been successfully submitted.</div>
                        <div class="alert alert-info" tkey="DealerFailureMessage" id="DealerFailureMessage" style="display:none">Your Submission is not successfully submitted.Please contact support team.</div>
                        <div class="alert alert-info" tkey="NoPDFMessage" id="NoPDFMessage" style="display:none">No PDF found.</div>
 
                        <div wfd-id="9" class="filter_area">
                            <div class="form-group row " style="display: flex;justify-content: space-between;flex-direction: row;">
                                <div class="row" style="width: 100%; max-width:1200px;padding-left:20px;">
                                    <div class="col-4" style = "display:contents;">
                                        <fieldset>
                                            <div class="control-group" style="display:flex; align-items: center;padding-top: 6px;">
 
                                                <label for="Dealer" class="form-label dc_area" tkey="DealerCode" style="display: flex; align-items: center;padding-left: 7px;">Clubhouse Code:</label>
                                                <input type="text"  style="width: 92px;height: 27px;" class="form-control dc_area" id="dealercode" placeholder="Clubhouse Code" required>
                                                <input type="hidden" id="userName" />
                                                <input type="hidden" id="dealercodehidden" />
                                            </div>
                                           
 
                                        </fieldset>
                                    </div>
                                    <div class="col-7">
                                        <fieldset>
                                            <div class="control-group" id="submissionweekdrp" style="display: flex;align-items: center;">
                                                <label for="Dealer" class="form-label dc_area" tkey="SubmissionWeek" style="display: flex; align-items: center;padding-right: 10px;">Submission Week:</label>
                                                <select style="width: 158px;display: flex;align-items: center;height: 28px;" class="form-control dc_area " id="submissionweek" placeholder="Submission Week"><option value=""></option></select>
 
                                                <button id="dealerFind" style="margin-left:23px" type="button" tkey="Find" class="btn btn-dark primary_btn">Find</button>
                                            </div>
                                        </fieldset>
                                    </div>
                                </div>
                                <div class="row" style="width: 100%; max-width:1200px;padding-left:20px;">
                                    <div class="col-11">
                                        <fieldset>
                                           
                                            <div class="control-group" id="submissionweektxt" style="padding-top:20px;">
 
                                                <label for="CurrentWeek" class="form-label dc_area text-left" tkey="CurrentWeekLabel">Current Week:</label>
 
                                                <label id="currentWeek" class="form-label  dc_area"> </label>
                                            </div>
 
                                        </fieldset>
                                    </div>
                                    </div>
 
                                </div>
                            &nbsp; &nbsp;&nbsp;
                            <div style="width:100%; margin:0 auto;">
                                <table id="dealers" class="table table-striped table-bordered dt-responsive nowrap data-entry" cellspacing="0"> </table>
                            </div>
                            <p style="text-align:right">
                                <button id="dealersubmitbtn" tkey="Submit" type="submit" value="Submit" style="display:none" class="btn btn-dark primary_btn">Submit</button>
                                <button id="dealerreset" tkey="Reset" class="btn btn-outline-secondary secondary_btn" style="display:none">Reset</button>
                            </p>
                        </div>
                    </div>
                </div>
                <div id="Summaries">
                    <h1 class="heading" tkey="TrafficClosingStatistics" id="TrafficClosingStatistics">Traffic & Closing Statistics</h1>
 
                    &nbsp; &nbsp;&nbsp;
 
                    <div id="DealerDetailReport" style="width:100%">
                        <div class="form-group row ">
                            <div class="row" style="width: 100%; padding-left:20px;">
                                <div class="col-4">
                                    <fieldset>
                                        <div class="control-group">
 
                                            <label for="DealerReport" class="form-label dc_area"   tkey="DealerCode">Clubhouse Code:</label>
                                            <input type="text"  style="width:140px !important" class="form-control" id="dealercodereport" placeholder="Dealer Code" readonly="readonly">
 
                                        </div>
                                       
                                    </fieldset>
                                </div>
                                <div class="col">
                                    <fieldset>
                                        <div class="control-group" style="text-align: right;  ">
                                            <button type="button" id="btnMTD" class="btn btn-dark primary_btn" style="background: rgb(204, 204, 204); margin-bottom: 5px">MTD</button>
                                            <button type="button" id="btnMTDTrend" class="btn btn-dark primary_btn" style="margin-bottom:10px">MTD Trend</button>
                                            <button type="button" id="btnWOW" class="btn btn-dark primary_btn">WOW</button>
                                            <button type="button" id="btnMOM" class="btn btn-dark primary_btn">MOM</button>
                                        </div>
                                       
                                    </fieldset>
                                </div>
                            </div>
                            <div class="row" style="width: 100%; padding-left:20px;">
                                <div class="col">
                                    <fieldset>
                                       
                                        <div class="control-group" style="padding-top:20px; display: flex;">
                                            <label for="ReportingWeek" class="form-label dc_area" style="padding-left:0px" tkey="ReportingWeek">Reporting Week:</label>
                                            <select style="width:150px" class="form-control dc_area " id="submissionReportweek" placeholder="Reporting Week"><option value=""></option></select>
                                            <!--<label id="currentWeekdealerreport" class="col-form-label dc_area"> </label>-->
 
                                        </div>
 
                                    </fieldset>
                                </div>
                                <div class="col">
                                    <fieldset>
                                         
                                        <div class="control-group" style="text-align: right; padding-top: 20px;  ">
                                            <button type="button" class="btn btn-dark primary_btn" tkey="ExportToCSV" id="export">Export </button>
                                        </div>
                                    </fieldset>
                                </div>
                            </div>
 
                        </div>
                     
                    </div>
                   
                    <div id="MTDReport" style="width:100%; overflow-wrap: break-word; margin-top: 21px;">
                        <table id="MTDSubmission" class="table table-striped table-bordered dt-responsive nowrap data-entry" cellspacing="0" style= "position: relative;height: 482.222px;width: 102.222%;">
                            <thead>
                                <tr><th rowspan="2" tkey="Model">Model</th><th colspan="15">MTD</th></tr>
                                <tr><th colspan="7" tkey="CurrentWeek">Current Week</th><th colspan="9" tkey="CurrentMonthMTD">Current Month (MTD)</th></tr>
                                <tr>
                                    <th key="Traffic" tkey="Traffic">Traffic </th>
                                    <th key="Writes" tkey="Writes"> Writes</th>
                                    <th key="Closing" tkey="Closing"> Closing%</th>
                                    <th tkey="MonthlySalesForecast">Monthly Sales Forecast</th>
                                    <th key="AvgTraffic" tkey="AvgTraffic" class="Avgitalic"> Avg. Traffic</th>
                                    <th key="AvgWrites" tkey="AvgWrites" class="Avgitalic"> Avg. Writes</th>
                                    <th key="AvgClosing" tkey="AvgClosing" class="Avgitalic"> Avg. Closing%</th>
                                    <th key="Traffic" tkey="Traffic"> Traffic </th>
                                    <th key="Writes" tkey="Writes"> Writes</th>
                                    <th key="Closing" tkey="Closing"> Closing%</th>
                                    <th tkey="MonthlySalesTarget">Monthly Sales Target</th>
                                    <th tkey="Achievenment">% Achievement</th>
                                    <th key="AvgTraffic" tkey="AvgTraffic" class="Avgitalic">Avg. Traffic</th>
                                    <th key="AvgWrites" tkey="AvgWrites" class="Avgitalic">Avg. Writes</th>
                                    <th key="AvgClosing" tkey="AvgClosing" class="Avgitalic">Avg. Closing%</th>
                                </tr>
                            </thead>
                            <tbody></tbody>
                        </table>
                    </div>
                    <div id="MTDTrendReport" style="width:100%; overflow-wrap: break-word; margin-top: 21px;">
                        <table id="MTDTrend" class="table table-striped table-bordered dt-responsive nowrap data-entry" cellspacing="0" style="display:none;">
                            <thead>
                                <tr><th rowspan="2" tkey="Model">Model</th><th colspan="20" tkey="DealerSubmissionText">District Summary</th></tr>
                                <tr><th colspan="4" tkey="Week1">Week 1</th><th colspan="4" tkey="Week2">Week 2</th><th colspan="4" tkey="Week3">Week 3</th><th colspan="4" tkey="Week4">Week 4</th><th colspan="4" tkey="Week5">Week 5</th></tr>
                                <tr>
                                    <th tkey="Traffic"> Traffic </th>
                                    <th tkey="Writes"> Writes</th>
                                    <th tkey="Closing"> Closing%</th>
                                    <th tkey="MonthlySalesForecast">Monthly Sales Forecast</th>
                                    <th tkey="Traffic"> Traffic </th>
                                    <th tkey="Writes"> Writes</th>
                                    <th tkey="Closing"> Closing%</th>
                                    <th tkey="MonthlySalesForecast">Monthly Sales Forecast</th>
                                    <th tkey="Traffic"> Traffic </th>
                                    <th tkey="Writes"> Writes</th>
                                    <th tkey="Closing"> Closing%</th>
                                    <th tkey="MonthlySalesForecast">Monthly Sales Forecast</th>
                                    <th tkey="Traffic"> Traffic </th>
                                    <th tkey="Writes"> Writes</th>
                                    <th tkey="Closing"> Closing%</th>
                                    <th tkey="MonthlySalesForecast">Monthly Sales Forecast</th>
                                    <th tkey="Traffic"> Traffic </th>
                                    <th tkey="Writes"> Writes</th>
                                    <th tkey="Closing"> Closing%</th>
                                    <th tkey="MonthlySalesForecast">Monthly Sales Forecast</th>
                                </tr>
                            </thead>
                            <tbody></tbody>
                        </table>
                    </div>
                    <div id="WOWReport" style="width:100%; overflow-wrap: break-word; margin-top: 21px;">
                        <table id="WOWSubmission" class="table table-striped table-bordered dt-responsive nowrap data-entry" cellspacing="0" style="display:none;">
                            <thead>
                                <tr><th rowspan="2" tkey="Model">Model</th><th colspan="7" tkey="WoWCOMPARISON">WoW COMPARISON</th><th colspan="6" tkey="WOWYoYCOMPARISON">Current Week YoY (Yeaer-over-Year) COMPARISON</th></tr>
                                <tr>
                                    <th tkey="DealerTrafficPercent">Clubhouse Traffic %</th>
                                    <th tkey="AreaAvgTrafficPercent" class="Avgitalic">Area Avg. Traffic %</th>
                                    <th tkey="DealerWritesPercent">Clubhouse Writes %</th>
                                    <th tkey="AreaAvgWrites" class="Avgitalic">Area Avg. Writes %</th>
                                    <th tkey="DealerClosing">Clubhouse Closing%</th>
                                    <th tkey="AreaAvgClosing" class="Avgitalic">Area Avg. Closing%</th>
                                    <th tkey="MonthlySalesForecastPercent">Monthly Sales Forecast %</th>
                                    <th tkey="DealerTrafficPercent">Clubhouse Traffic %</th>
                                    <th tkey="AreaAvgTrafficPercent" class="Avgitalic">Area Avg. Traffic %</th>
                                    <th tkey="DealerWritesPercent">Clubhouse Writes %</th>
                                    <th tkey="AreaAvgWritesPercent" class="Avgitalic">Area Avg. Writes %</th>
                                    <th tkey="DealerClosing">Clubhouse Closing%</th>
                                    <th tkey="AreaAvgClosing" class="Avgitalic">Area Avg. Closing%</th>
                                </tr>
                            </thead>
                            <tbody></tbody>
                        </table>
                    </div>
                    <div id="MOMReport" style="width:100%; overflow-wrap: break-word; margin-top: 21px;">
                        <table id="MOMSubmission" class="table table-striped table-bordered dt-responsive nowrap data-entry" cellspacing="0" style="display:none;">
                            <thead>
                                <!--<tr><th rowspan="2" tkey="Model">Model</th><th colspan="7" tkey="MoMCOMPARISON">MOM COMPARISON</th><th colspan="6" tkey="YoYCOMPARISON">Current MTD YoY (Year-over-Year) COMPARISON</th></tr>-->
                                <tr><th rowspan="2" tkey="Model">Model</th><th colspan="7" tkey="MoMCOMPARISON">MOM COMPARISON</th><th colspan="6" tkey="YoYCOMPARISONDetailMOM">Current MTD YoY (Year-over-Year) COMPARISON</th></tr>
                                <tr>
                                    <th tkey="DealerTrafficPercent">Clubhouse Traffic % </th>
                                    <th tkey="AreaAvgTrafficPercent" class="Avgitalic">Area Avg. Traffic %</th>
                                    <th tkey="DealerWritesPercent">Clubhouse Writes %</th>
                                    <th tkey="AreaAvgWrites" class="Avgitalic">Area Avg. Writes %</th>
                                    <th tkey="DealerClosing">Clubhouse Closing%</th>
                                    <th tkey="AreaAvgClosing" class="Avgitalic">Area Avg. Closing%</th>
                                    <th tkey="MonthlySalesForecastPercent">Monthly Sales Forecast %</th>
                                    <th tkey="DealerTrafficPercent">Clubhouse Traffic %</th>
                                    <th tkey="AreaAvgTraffic" class="Avgitalic">Area Avg. Traffic %</th>
                                    <th tkey="DealerWritesPercent">Clubhouse Writes %</th>
                                    <th tkey="AreaAvgWritesPercent" class="Avgitalic">Area Avg. Writes %</th>
                                    <th tkey="DealerClosing">Clubhouse Closing%</th>
                                    <th tkey="AreaAvgClosing" class="Avgitalic">Area Avg. Closing%</th>
                                </tr>
                            </thead>
                            <tbody></tbody>
                        </table>
                    </div>
                </div>
            </div>
        </div>
    </div>
 
`;
        this.domElement.innerHTML = tabsHTML;
 
        // Initialize jQuery tabs include your tab function here add by Sagnik
       
        (jQuery("#tabs", this.domElement) as any).tabs();
        let culture = this.context.pageContext.cultureInfo.currentUICultureName;   
        let requestUrl = "";
        if(culture.toString()==='en-US')
        {
            requestUrl=this.context.pageContext.web.absoluteUrl.concat("/_api/web/lists/getbytitle('TrafficLog Documents')/items?$select=ID,Title,ServerRedirectedEmbedUri,Modified&$filter=Language eq 'English'&$orderby=Modified desc");
        }
        else
        {
            requestUrl=this.context.pageContext.web.absoluteUrl.concat("/_api/web/lists/getbytitle('TrafficLog Documents')/items?$select=ID,Title,ServerRedirectedEmbedUri,Modified&$filter=Language eq 'French'&$orderby=Modified desc");
        } 
        debugger;
    this.context.spHttpClient.get(requestUrl, SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) => {
            if (response.ok) {
               response.json().then((responseJSON) => {
                    if (responseJSON!=null && responseJSON.value!=null){
            //let itemCount:number = parseInt(responseJSON.value.toString());
            //console.log(itemCount);
            var newdiv = document.createElement("div");
            newdiv.innerHTML = '<a href="'+responseJSON.value[0].ServerRedirectedEmbedUri+'" target="_blank" data-interception="off">Download Traffic Calendar (PDF)</a>';
            //var pdfUrl = responseJSON.value[0].ServerRedirectedEmbedUri;
            //newdiv.innerHTML = '<a href="#" onclick="window.open('+pdfUrl+');return false;">Download Traffic Calendar (PDF)</a>';
            //this.domElement.append(newdiv);
             document.getElementById("divDownload")?.append(newdiv);
            return false;
                    }
                });
            }
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