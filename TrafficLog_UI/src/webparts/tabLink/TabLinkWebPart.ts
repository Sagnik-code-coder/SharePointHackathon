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

import * as strings from 'TabLinkWebPartStrings';
import TabLink from './components/TabLink';
import { ITabLinkProps } from './components/ITabLinkProps';
import * as jQuery from 'jquery';
import 'jqueryui';


export interface ITabLinkWebPartProps {
  description: string;
}

export default class TabLinkWebPart extends BaseClientSideWebPart<ITabLinkWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<ITabLinkProps> = React.createElement(
      TabLink,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
    const tabsHTML: string =
      `
    <style type="text/css">
        .ms-webpart-titleText {
            display: none;
        }
    </style>
    <script type="text/javascript">
         document.onreadystatechange = function () {
             if (document.readyState !== "complete") {
                 document.querySelector(".body-content").style.visibility = "hidden";
                 document.querySelector( "#loader").style.visibility = "visible";
             } else {
                 document.querySelector("#loader").style.display = "none";
                 document.querySelector(".body-content").style.visibility = "visible";
             }
         };
    </script> 
</head>
<body>
    <div id="loader" class="centerloader"></div>
    <div class="body-content">
        <div class="container-fluid">
            <div id="tabs" style="">
            <span style="display: none;" id="Currculture">`+ this.context.pageContext.cultureInfo.currentCultureName +`</span>
                    <ul>
                        <li><a href="#TrafficDetail" tkey="TrafficDetail">Traffic Detail</a></li>
                        <li><a href="#TrafficSummary" tkey="TrafficSummary">Traffic Summary</a></li>
                    </ul>
                    <div id="TrafficDetail">
                
                        <div class="alert alert-info" tkey="UserInaccessibility" id="UserInaccessibilityAlert" style="display:none">The current user is not allowed to visit this page.</div>
                        
                        <div id="detail">
                            <div id="dealersummary">
                                <div class="row" style="width: 100%;max-width:2000px; ">
                                    <div class="span4">
                                        <fieldset>
                                            <div class="control-group">
        
                                                <label for="RegionArea"  style="padding-left:15px" class="form-label dc_area" tkey="RegionArea">Region/Area/Volume Group:</label>
                                                 <select style="width:140px !important" class="form-control  dc_area" id="Regions"></select>
                                            </div>
        
                                        </fieldset>
                                    </div>
                                    <div class="span3">
                                        <fieldset>
                                            <div class="control-group">
                                                <label for="DistrictArea" class="form-label dc_area" tkey="DistrictArea">District:</label>
                                                <select style="width:140px !important" class="form-control dc_area" id="Districts"></select>
                                            </div>
                                        </fieldset>
                                    </div>
                                    <div class="span3">
                                        <fieldset>
                                            <div class="control-group">
                                                <label for="Dealers" class="form-label dc_area" tkey="Dealers">Clubhouses:</label>
                                                <select style="width:140px !important" class="form-control dc_area" id="Dealers"></select>
                                            </div>
                                        </fieldset>
                                    </div>
                                </div>
                                              
                                <div class="row" style="width: 100%; padding-top:20px;">
         
                                        <div class="col">
                                            <fieldset>
                                                <div class="control-group"  id="TrafficDetailReport" style="text-align: right;">
                                                    <button type="button" id="btnMTD" class="btn btn-dark primary_btn">MTD</button>
                                                    <button type="button" id="btnMTDTrend" class="btn btn-dark primary_btn">MTD Trend</button>
                                                    <button type="button" id="btnWOW" class="btn btn-dark primary_btn">WOW</button>
                                                    <button   type="button" id="btnMOM" class="btn btn-dark primary_btn">MOM</button>
                                                </div>
                                                
                                            </fieldset>
                                        </div>
                                    </div>
        
                                <div class="row" style="width: 100%; padding-top:20px;">
                                    <div class="col">
                                        <fieldset>
        
                                            <div class="control-group">
                                                <label for="CurrentWeek" class="form-label dc_area" tkey="CurrentWeekLabel">Current Week:</label>
                                                
                                                <label id="currentWeek" class="form-label dc_area"> </label>
        
                                            </div>
        
                                        </fieldset>
                                    </div>
                                    <div class="col">
                                        <fieldset>
                                            
                                            <div class="control-group" style="text-align: right;">
                                                <button type="button" class="btn btn-dark primary_btn" tkey="ExportToCSV" id="exportdealer">Export</button>
                                            </div>
                                        </fieldset>
                                    </div>
                                </div>
           
                                
        
        
                                
                               
                                    <div id="MTDReport" style="width:100%">
                                        <table id="MTDSubmission" class="table table-striped table-bordered dt-responsive nowrap data-entry" cellspacing="0" style="">
                                            <thead>
                                                <tr><th rowspan="3" tkey="Model">Model</th><th colspan="15">MTD</th></tr>
                                                <tr><th colspan="7" tkey="CurrentWeek">Current Week</th><th colspan="9" tkey="CurrentMonthMTD">Current Month (MTD)</th></tr>
                                                <tr>
                                                    <th key="Traffic" tkey="Traffic">Traffic </th>
                                                    <th key="Writes" tkey="Writes"> Writes</th>
                                                    <th key="Closing" tkey="Closing"> Closing%</th>
                                                    <th tkey="MonthlySalesForecast">Monthly Sales Forecast</th>
                                                    <th key="AvgTraffic" tkey="AvgTraffic"> Avg. Traffic</th>
                                                    <th key="AvgWrites" tkey="AvgWrites"> Avg. Writes</th>
                                                    <th key="AvgClosing" tkey="AvgClosing"> Avg. Closing%</th>
                                                    <th key="Traffic" tkey="Traffic"> Traffic </th>
                                                    <th key="Writes" tkey="Writes"> Writes</th>
                                                    <th key="Closing" tkey="Closing"> Closing%</th>
                                                    <th tkey="MonthlySalesTarget">Monthly Sales Target</th>
                                                    <th tkey="Achievenment">% Achievement</th>
                                                    <th key="AvgTraffic" tkey="AvgTraffic">Avg. Traffic</th>
                                                    <th key="AvgWrites" tkey="AvgWrites">Avg. Writes</th>
                                                    <th key="AvgClosing" tkey="AvgClosing">Avg. Closing%</th>
                                                </tr>
                                            </thead>
                                            <tbody></tbody>
                                        </table>
                                    </div>
                                    <div id="MTDTrendReport" style="width:100%">
                                        <table id="MTDTrend" class="table table-striped table-bordered dt-responsive nowrap data-entry" cellspacing="0" style="display:none;">
                                            <thead>
                                                <tr><th rowspan="7" tkey="Model">Model</th><th colspan="20" tkey="MTDTrend">MTD Trend</th></tr>
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
                                    <div id="WOWReport" style="width:100%">
                                        <table id="WOWSubmission" class="table table-striped table-bordered dt-responsive nowrap data-entry" cellspacing="0" style="display:none;">
                                            <thead>
                                                <tr><th rowspan="2" tkey="Model">Model</th><th colspan="7" tkey="WoWCOMPARISON">WoW COMPARISON</th><th colspan="6" tkey="WOWYoYCOMPARISON">Current Week YoY COMPARISON</th></tr>
                                                <tr>
                                                    <th tkey="DealerTrafficPercent">Clubhouse Traffic %</th>
                                                    <th tkey="AreaAvgTrafficPercent">Area Avg. Traffic %</th>
                                                    <th tkey="DealerWritesPercent">Clubhouse Writes %</th>
                                                    <th tkey="AreaAvgWrites">Area Avg. Writes %</th>
                                                    <th tkey="DealerClosing">Clubhouse Closing (ppt)</th>
                                                    <th tkey="AreaAvgClosing">Area Avg. Closing (ppt)</th>
                                                    <th tkey="MonthlySalesForecast">Monthly Sales Forecast %</th>
                                                    <th tkey="DealerTrafficPercent">Clubhouse Traffic %</th>
                                                    <th tkey="AreaAvgTrafficPercent">Area Avg. Traffic %</th>
                                                    <th tkey="DealerWritesPercent">Clubhouse Writes %</th>
                                                    <th tkey="AreaAvgWritesPercent">Area Avg. Writes %</th>
                                                    <th tkey="DealerClosing">Clubhouse Closing (ppt)</th>
                                                    <th tkey="AreaAvgClosing">Area Avg. Closing (ppt)</th>
                                                </tr>
                                            </thead>
                                            <tbody></tbody>
                                        </table>
                                    </div>
                                    <div id="MOMReport" style="width:100%">
                                        <table id="MOMSubmission" class="table table-striped table-bordered dt-responsive nowrap data-entry" cellspacing="0" style="display:none;">
                                            <thead>
                                                <!--<tr><th rowspan="2" tkey="Model">Model</th><th colspan="7" tkey="MoMCOMPARISON">MOM COMPARISON</th><th colspan="6" tkey="YoYCOMPARISON">MTD YoY COMPARISON</th></tr>-->
                                                <tr><th rowspan="2" tkey="Model">Model</th><th colspan="7" tkey="MoMCOMPARISON">MOM COMPARISON</th><th colspan="6" tkey="YoYCOMPARISONDetailMOM">Current MTD YoY (Year-over-Year) COMPARISON</th></tr>
                                                <tr>
                                                    <th tkey="DealerTrafficPercent">Clubhouse Traffic % </th>
                                                    <th tkey="AreaAvgTrafficPercent">Area Avg. Traffic %</th>
                                                    <th tkey="DealerWritesPercent">Clubhouse Writes %</th>
                                                    <th tkey="AreaAvgWrites">Area Avg. Writes %</th>
                                                    <th tkey="DealerClosing">Clubhouse Closing (ppt)</th>
                                                    <th tkey="AreaAvgClosing">Area Avg. Closing (ppt)</th>
                                                    <th tkey="MonthlySalesForecastPercent">Monthly Sales Forecast %</th>
                                                    <th tkey="DealerTrafficPercent">Clubhouse Traffic %</th>
                                                    <th tkey="AreaAvgTraffic">Area Avg. Traffic %</th>
                                                    <th tkey="DealerWritesPercent">Clubhouse Writes %</th>
                                                    <th tkey="AreaAvgWritesPercent">Area Avg. Writes %</th>
                                                    <th tkey="DealerClosing">Clubhouse Closing (ppt)</th>
                                                    <th tkey="AreaAvgClosing">Area Avg. Closing (ppt)</th>
                                                </tr>
                                            </thead>
                                            <tbody></tbody>
                                        </table>
                                    </div>
        
                                </div>
                        </div>
        
                    </div>
                    <div id="TrafficSummary" style="display:none">

                        <div class="alert alert-info" tkey="UserInaccessibility" id="UserInaccessibility" style="display:none">The current user is not allowed to visit this page.</div>
        
                        <div id="summary">
                            <div id="dealersummary">
                                <div class="row" style="width: 100%;max-width:2000px; ">
                                    <div class="span4">
                                        <fieldset>
                                            <div class="control-group">
        
        
        
                                                <label for="RegionArea" style="padding-left:15px" class="form-label dc_area" tkey="RegionAreaSummary">Region:</label>
                                                <select style="width:140px !important" class="form-control dc_area" id="RegionsSummary"></select>
                                            </div>
        
                                        </fieldset>
                                    </div>
                                    <div class="span3">
                                        <fieldset>
                                            <div class="control-group">
        
                                                <label for="DistrictArea" class="form-label dc_area" tkey="DistrictArea">District:</label>
                                                <select style="width:140px !important" class="form-control dc_area" id="DistrictsSummary"></select>
        
                                            </div>
                                        </fieldset>
                                    </div>
                                    <div class="span3">
                                        <fieldset>
                                            <div class="control-group" style="display:none;">
        
                                                <label for="Dealers" class="form-label dc_area" tkey="Dealers">Clubhouse:</label>
                                                <select class="form-control dc_area" id="DealersSummary"></select>
                                            </div>
                                        </fieldset>
                                    </div>
                                </div>
        
                                <div class="row" style="width: 100%; padding-top:20px;">
         
                                    <div class="col">
                                        <fieldset>
                                            <div class="control-group" id="TrafficSummaryReport" style="text-align: right;">
                                                <button type="button" id="btnMTDSummary" class="btn btn-dark primary_btn">MTD</button>
                                                <button type="button" id="btnMTDTrendSummary" class="btn btn-dark primary_btn">MTD Trend</button>
                                                <button type="button" id="btnWOWSummary" class="btn btn-dark primary_btn">WOW</button>
                                                <button type="button" id="btnMOMSummary" class="btn btn-dark primary_btn">MOM</button>
                                            </div>
        
                                        </fieldset>
                                    </div>
                                </div>
        
                                <div class="row" style="width: 100%;  padding-top:20px;">
                                    <div class="col">
                                        <fieldset>
        
                                            <div class="control-group">
                                               
                                                <label for="CurrentWeek" class="form-label dc_area" tkey="CurrentWeekLabel">Current Week:</label>
        
                                                <label id="currentWeekSummary" class="col-form-label dc_area"> </label>
                                            </div>
        
                                        </fieldset>
                                    </div>
                                    <div class="col">
                                        <fieldset>
        
                                            <div class="control-group" style="text-align: right;">
                                                <button type="button" class="btn btn-dark primary_btn" tkey="ExportToCSV" id="exportSummary">Export</button>
         
                                            </div>
                                        </fieldset>
                                    </div>
                                </div>
                              
                                
                                
                                <div id="MTDReport" style="width:100%">
                                    <table id="MTDDistrict" class="table table-striped table-bordered dt-responsive nowrap data-entry" cellspacing="0" style="display:none;">
                                        <thead>
                                            <tr><th rowspan="6" tkey="DealerNumber">Area</th>
                                                            <th colspan="15" tkey="DistrictDealerSummaryText">National Area Summary</th></tr>
                                            <tr><th colspan="7" tkey="CurrentWeek">Current Week</th><th colspan="9" tkey="CurrentMonthMTD">Current Month (MTD)</th></tr>
                                            <tr>
                                                <th tkey="DealerTraffic">Traffic </th>
                                                <th tkey="DealerWrites">Writes</th>
                                                <th tkey="DealerClosing">Closing%</th>
                                                <th tkey="MonthlySalesForecast">Monthly Sales Forecast</th>
                                                <th tkey="AreaAvgTraffic">Area Avg. Traffic</th>
                                                <th tkey="AreaAvgWrites">Area Avg. Writes</th>
                                                <th tkey="AreaAvgClosing">Area Avg. Closing%</th>
                                                <th tkey="DealerTraffic">Traffic </th>
                                                <th tkey="DealerWrites">Writes</th>
                                                <th tkey="DealerClosing">Closing%</th>
                                                <th tkey="MonthlySalesTarget">Monthly Sales Target</th>
                                                <th tkey="Achievenment">% Achievement</th>
                                                <th tkey="AreaAvgTraffic"> Area Avg. Traffic</th>
                                                <th tkey="AreaAvgWrites">Area Avg. Writes</th>
                                                <th tkey="AreaAvgClosing">Area Avg. Closing%</th>
                                            </tr>
                                        </thead>
                                        <tbody></tbody>
                                    </table>
                                    <table id="MTDRegionDistrict" class="table table-striped table-bordered dt-responsive nowrap data-entry" cellspacing="0" style="display:none;">
                                        <thead>
                                            <tr><th rowspan="3" tkey="DealerDistrict"></th><th colspan="15" tkey="RegionDistrictSummaryText">Region Summary</th></tr>
                                            <tr><th colspan="7" tkey="CurrentWeek">Current Week</th><th colspan="9" tkey="CurrentMonthMTD">Current Month (MTD)</th></tr>
                                            <tr>
                                                <th tkey="Traffic">Traffic</th>
                                                <th tkey="Writes">Writes</th>
                                                <th tkey="Closing">Closing%</th>
                                                <th tkey="MonthlySalesForecast">Monthly Sales Forecast</th>
                                                <th tkey="NationalAvgTraffic">National Avg. Traffic</th>
                                                <th tkey="NationalAvgWrites">National Avg. Writes</th>
                                                <th tkey="NationalAvgClosing">National Avg. Closing%</th>
                                                <th tkey="Traffic">Traffic </th>
                                                <th tkey="Writes">Writes</th>
                                                <th tkey="Closing">Closing%</th>
                                                <th tkey="MonthlySalesTarget">Monthly Sales Target</th>
                                                <th tkey="Achievenment">% Achievement</th>
                                                <th tkey="NationalAvgTraffic"> National Avg. Traffic</th>
                                                <th tkey="NationalAvgWrites">National Avg. Writes</th>
                                                <th tkey="NationalAvgClosing">National Avg. Closing%</th>
                                            </tr>
                                        </thead>
                                        <tbody></tbody>
                                    </table>
                                    <table id="MTDNationalArea" class="table table-striped table-bordered dt-responsive nowrap data-entry" cellspacing="0" style="display:none;">
                                        <thead>
                                            <tr><th rowspan="3" tkey="DealerArea"></th><th colspan="9" tkey="NationalAreaSummaryText">National Area Summary</th></tr>
                                            <tr><th colspan="4" tkey="CurrentWeek">Current Week</th><th colspan="5" tkey="CurrentMonthMTD">Current Month (MTD)</th></tr>
                                            <tr>
                                                <th tkey="Traffic">Traffic </th>
                                                <th tkey="Writes">Writes</th>
                                                <th tkey="Closing">Closing%</th>
                                                <th tkey="MonthlySalesForecast">Monthly Sales Forecast</th>
                                                <th tkey="Traffic">Traffic </th>
                                                <th tkey="Writes">Writes</th>
                                                <th tkey="Closing">Closing%</th>
                                                <th tkey="MonthlySalesTarget">Monthly Sales Target</th>
                                                <th tkey="Achievenment">% Achievement</th>
                                            </tr>
                                        </thead>
                                        <tbody></tbody>
                                    </table>
                                    <table id="MTDNationalRegion" class="table table-striped table-bordered dt-responsive nowrap data-entry" cellspacing="0" style="display:none;">
                                        <thead>
                                            <tr><th rowspan="3" tkey="DealerRegion"></th><th colspan="9" tkey="NationalRegionSummaryText">National Region Summary</th></tr>
                                            <tr><th colspan="4" tkey="CurrentWeek">Current Week</th><th colspan="5" tkey="CurrentMonthMTD">Current Month (MTD)</th></tr>
                                            <tr>
                                                <th tkey="Traffic">Traffic </th>
                                                <th tkey="Writes">Writes</th>
                                                <th tkey="Closing">Closing%</th>
                                                <th tkey="MonthlySalesForecast">Monthly Sales Forecast</th>
                                                <th tkey="Traffic">Traffic </th>
                                                <th tkey="Writes">Writes</th>
                                                <th tkey="Closing">Closing%</th>
                                                <th tkey="MonthlySalesTarget">Monthly Sales Target</th>
                                                <th tkey="Achievenment">% Achievement</th>
                                            </tr>
                                        </thead>
                                        <tbody></tbody>
                                    </table>
                                </div>
                                <div id="MTDTrendReport" style="width:100%">
                                    <table id="MTDTrendDistrict" class="table table-striped table-bordered dt-responsive nowrap data-entry" cellspacing="0" style="display:none;">
                                        <thead>
                                            <tr><th rowspan="3" tkey="DealerNumber">Area</th><th colspan="20" tkey="DistrictDealerSummaryText">National Area Summary</th></tr>
                                            <tr><th colspan="4" tkey="Week1">Week 1</th><th colspan="4" tkey="Week2">Week 2</th><th colspan="4" tkey="Week3">Week 3</th><th colspan="4" tkey="Week4">Week 4</th><th colspan="4" tkey="Week5">Week 5 - Current Week</th></tr>
                                            <tr>
                                                <th tkey="DealerTraffic">Traffic </th>
                                                <th tkey="DealerWrites">Writes</th>
                                                <th tkey="DealerClosing">Closing%</th>
                                                <th tkey="MonthlySalesForecast">Monthly Sales Forecast</th>
                                                <th tkey="DealerTraffic">Traffic </th>
                                                <th tkey="DealerWrites">Writes</th>
                                                <th tkey="DealerClosing">Closing%</th>
                                                <th tkey="MonthlySalesForecast">Monthly Sales Forecast</th>
                                                <th tkey="DealerTraffic">Traffic </th>
                                                <th tkey="DealerWrites">Writes</th>
                                                <th tkey="DealerClosing">Closing%</th>
                                                <th tkey="MonthlySalesForecast">Monthly Sales Forecast</th>
                                                <th tkey="DealerTraffic">Traffic </th>
                                                <th tkey="DealerWrites">Writes</th>
                                                <th tkey="DealerClosing">Closing%</th>
                                                <th tkey="MonthlySalesForecast">Monthly Sales Forecast</th>
                                                <th tkey="DealerTraffic">Traffic </th>
                                                <th tkey="DealerWrites">Writes</th>
                                                <th tkey="DealerClosing">Closing%</th>
                                                <th tkey="MonthlySalesForecast">Monthly Sales Forecast</th>
                                            </tr>
                                        </thead>
                                        <tbody></tbody>
                                    </table>
                                    <table id="MTDTrendRegionDistrict" class="table table-striped table-bordered dt-responsive nowrap data-entry" cellspacing="0" style="display:none;">
                                        <thead>
                                            <tr><th rowspan="3" tkey="DealerDistrict"></th><th colspan="20" tkey="RegionDistrictSummaryText">Region Summary</th></tr>
                                            <tr><th colspan="4" tkey="Week1">Week 1</th><th colspan="4" tkey="Week2">Week 2</th><th colspan="4" tkey="Week3">Week 3</th><th colspan="4" tkey="Week4">Week 4</th><th colspan="4" tkey="Week5">Week 5 - Current Week</th></tr>
                                            <tr>
                                                <th tkey="DealerTraffic">Clubhouse Traffic </th>
                                                <th tkey="DealerWrites">Clubhouse Writes</th>
                                                <th tkey="DealerClosing">Clubhouse Closing%</th>
                                                <th tkey="MonthlySalesForecast">Monthly Sales Forecast</th>
                                                <th tkey="DealerTraffic">Clubhouse Traffic </th>
                                                <th tkey="DealerWrites">Clubhouse Writes</th>
                                                <th tkey="DealerClosing">Clubhouse Closing%</th>
                                                <th tkey="MonthlySalesForecast">Monthly Sales Forecast</th>
                                                <th tkey="DealerTraffic">Clubhouse Traffic </th>
                                                <th tkey="DealerWrites">Clubhouse Writes</th>
                                                <th tkey="DealerClosing">Clubhouse Closing%</th>
                                                <th tkey="MonthlySalesForecast">Monthly Sales Forecast</th>
                                                <th tkey="DealerTraffic">Clubhouse Traffic </th>
                                                <th tkey="DealerWrites">Clubhouse Writes</th>
                                                <th tkey="DealerClosing">Clubhouse Closing%</th>
                                                <th tkey="MonthlySalesForecast">Monthly Sales Forecast</th>
                                                <th tkey="DealerTraffic">Clubhouse Traffic </th>
                                                <th tkey="DealerWrites">Clubhouse Writes</th>
                                                <th tkey="DealerClosing">Clubhouse Closing%</th>
                                                <th tkey="MonthlySalesForecast">Monthly Sales Forecast</th>
                                            </tr>
                                        </thead>
                                        <tbody></tbody>
                                    </table>
                                    <table id="MTDTrendNationalArea" class="table table-striped table-bordered dt-responsive nowrap data-entry" cellspacing="0" style="display:none;">
                                        <thead>
                                            <tr><th rowspan="3" tkey="DealerArea"></th><th colspan="20" tkey="NationalAreaSummaryText">National Area Summary</th></tr>
                                            <tr><th colspan="4" tkey="Week1">Week 1</th><th colspan="4" tkey="Week2">Week 2</th><th colspan="4" tkey="Week3">Week 3</th><th colspan="4" tkey="Week4">Week 4</th><th colspan="4" tkey="Week5">Week 5 - Current Week</th></tr>
                                            <tr>
                                                <th tkey="DealerTraffic">Clubhouse Traffic </th>
                                                <th tkey="DealerWrites">Clubhouse Writes</th>
                                                <th tkey="DealerClosing">Clubhouse Closing%</th>
                                                <th tkey="MonthlySalesForecast">Monthly Sales Forecast</th>
                                                <th tkey="DealerTraffic">Clubhouse Traffic </th>
                                                <th tkey="DealerWrites">Clubhouse Writes</th>
                                                <th tkey="DealerClosing">Clubhouse Closing%</th>
                                                <th tkey="MonthlySalesForecast">Monthly Sales Forecast</th>
                                                <th tkey="DealerTraffic">Clubhouse Traffic </th>
                                                <th tkey="DealerWrites">Clubhouse Writes</th>
                                                <th tkey="DealerClosing">Clubhouse Closing%</th>
                                                <th tkey="MonthlySalesForecast">Monthly Sales Forecast</th>
                                                <th tkey="DealerTraffic">Clubhouse Traffic </th>
                                                <th tkey="DealerWrites">Clubhouse Writes</th>
                                                <th tkey="DealerClosing">Clubhouse Closing%</th>
                                                <th tkey="MonthlySalesForecast">Monthly Sales Forecast</th>
                                                <th tkey="DealerTraffic">Clubhouse Traffic </th>
                                                <th tkey="DealerWrites">Clubhouse Writes</th>
                                                <th tkey="DealerClosing">Clubhouse Closing%</th>
                                                <th tkey="MonthlySalesForecast">Monthly Sales Forecast</th>
                                            </tr>
                                        </thead>
                                        <tbody></tbody>
                                    </table>
                                    <table id="MTDTrendNationalRegion" class="table table-striped table-bordered dt-responsive nowrap data-entry" cellspacing="0" style="display:none;">
                                        <thead>
                                            <tr><th rowspan="3" tkey="DealerRegion"></th><th colspan="20" tkey="NationalRegionSummaryText">National Region Summary</th></tr>
                                            <tr><th colspan="4" tkey="Week1">Week 1</th><th colspan="4" tkey="Week2">Week 2</th><th colspan="4" tkey="Week3">Week 3</th><th colspan="4" tkey="Week4">Week 4</th><th colspan="4" tkey="Week5">Week 5 - Current Week</th></tr>
                                            <tr>
                                                <th tkey="DealerTraffic">Clubhouse Traffic </th>
                                                <th tkey="DealerWrites">Clubhouse Writes</th>
                                                <th tkey="DealerClosing">Clubhouse Closing%</th>
                                                <th tkey="MonthlySalesForecast">Monthly Sales Forecast</th>
                                                <th tkey="DealerTraffic">Clubhouse Traffic </th>
                                                <th tkey="DealerWrites">Clubhouse Writes</th>
                                                <th tkey="DealerClosing">Clubhouse Closing%</th>
                                                <th tkey="MonthlySalesForecast">Monthly Sales Forecast</th>
                                                <th tkey="DealerTraffic">Clubhouse Traffic </th>
                                                <th tkey="DealerWrites">Clubhouse Writes</th>
                                                <th tkey="DealerClosing">Clubhouse Closing%</th>
                                                <th tkey="MonthlySalesForecast">Monthly Sales Forecast</th>
                                                <th tkey="DealerTraffic">Clubhouse Traffic </th>
                                                <th tkey="DealerWrites">Clubhouse Writes</th>
                                                <th tkey="DealerClosing">Clubhouse Closing%</th>
                                                <th tkey="MonthlySalesForecast">Monthly Sales Forecast</th>
                                                <th tkey="DealerTraffic">Clubhouse Traffic </th>
                                                <th tkey="DealerWrites">Clubhouse Writes</th>
                                                <th tkey="DealerClosing">Clubhouse Closing%</th>
                                                <th tkey="MonthlySalesForecast">Monthly Sales Forecast</th>
                                            </tr>
                                        </thead>
                                        <tbody></tbody>
                                    </table>
                                </div>
                                <div id="WOWReport" style="width:100%">
                                    <table id="WOWDistrict" class="table table-striped table-bordered dt-responsive nowrap data-entry" cellspacing="0" style="display:none;">
                                        <thead>
                                            <tr><th rowspan="2" tkey="DealerNumber">Clubhouse Number</th><th rowspan="2" tkey="DealerName">Clubhouse Name</th><th colspan="7" tkey="WoWCOMPARISON">WoW COMPARISON</th><th colspan="6" tkey="YoYCOMPARISON">Current Week YoY (Year-over-Year) COMPARISON</th></tr>
                                            <tr>
                                                <th tkey="DealerTrafficPercent">Clubhouse Traffic </th>
                                                <th tkey="AreaAvgTrafficPercent">Area Avg. Traffic %</th>
                                                <th tkey="DealerWritesPercent">Clubhouse Writes</th>
                                                <th tkey="AreaAvgWritesPercent">Area Avg. Writes</th>
                                                <th tkey="DealerClosingNoPercent">Clubhouse Closing%</th>
                                                <th tkey="AreaAvgClosingNoPercent">Area Avg. Closing%</th>
                                                <th tkey="MonthlySalesForecastPercent">Monthly Sales Forecast</th>
                                                <th tkey="DealerTrafficPercent">Clubhouse Traffic</th>
                                                <th tkey="AreaAvgTrafficPercent">Area Avg. Traffic %</th>
                                                <th tkey="DealerWritesPercent">Clubhouse Writes</th>
                                                <th tkey="AreaAvgWritesPercent">Area Avg. Writes</th>
                                                <th tkey="DealerClosingNoPercent">Clubhouse Closing%</th>
                                                <th tkey="AreaAvgClosingNoPercent">Area Avg. Closing%</th>
                                            </tr>
                                        </thead>
                                        <tbody></tbody>
                                    </table>
                                    <table id="WOWRegionDistrict" class="table table-striped table-bordered dt-responsive nowrap data-entry" cellspacing="0" style="display:none;">
                                        <thead>
                                            <tr><th rowspan="2" tkey="DealerDistrict"></th><th colspan="7" tkey="WoWCOMPARISON">WoW COMPARISON</th><th colspan="6" tkey="YoYCOMPARISON">Current Week YoY (Year-over-Year) COMPARISON</th></tr>
                                            <tr>
                                                <th tkey="TrafficPerent">Traffic </th>
                                                <th tkey="NationalAvgTrafficPercent">National Avg. Traffic</th>
                                                <th tkey="WritesPercent">Writes</th>
                                                <th tkey="NationalAvgWritesPercent">National Avg. Writes</th>
                                                <th tkey="ClosingNoPercent">Closing%</th>
                                                <th tkey="NationalAvgClosingNoPercent">National Avg. Closing%</th>
                                                <th tkey="MonthlySalesForecastPercent">Monthly Sales Forecast</th>
                                                <th tkey="TrafficPerent">Traffic</th>
                                                <th tkey="NationalAvgTrafficPercent">National Avg. Traffic</th>
                                                <th tkey="WritesPercent">Writes</th>
                                                <th tkey="NationalAvgWritesPercent">National Avg. Writes</th>
                                                <th tkey="ClosingNoPercent">Closing%</th>
                                                <th tkey="NationalAvgClosingNoPercent">National Avg. Closing%</th>
                                            </tr>
                                        </thead>
                                        <tbody></tbody>
                                    </table>
                                    <table id="WOWNationalArea" class="table table-striped table-bordered dt-responsive nowrap data-entry" cellspacing="0" style="display:none;">
                                        <thead>
                                            <tr><th rowspan="2" tkey="DealerArea"></th><th colspan="4" tkey="WoWCOMPARISON">WoW COMPARISON</th><th colspan="3" tkey="YoYCOMPARISON">Current Week YoY (Year-over-Year) COMPARISON</th></tr>
                                            <tr>
                                                <th tkey="TrafficPerent">Traffic </th>
                                                <th tkey="WritesPercent">Writes</th>
                                                <th tkey="ClosingNoPercent">Closing%</th>
                                                <th tkey="MonthlySalesForecastPercent">Monthly Sales Forecast</th>
                                                <th tkey="TrafficPercent">Traffic</th>
                                                <th tkey="WritesPercent">Writes</th>
                                                <th tkey="ClosingNoPercent">Closing%</th>
                                            </tr>
                                        </thead>
                                        <tbody></tbody>
                                    </table>
                                    <table id="WOWNationalRegion" class="table table-striped table-bordered dt-responsive nowrap data-entry" cellspacing="0" style="display:none;">
                                        <thead>
                                            <tr><th rowspan="2" tkey="DealerRegion"></th><th colspan="4" tkey="WoWCOMPARISON">WoW COMPARISON</th><th colspan="3" tkey="YoYCOMPARISON">Current Week YoY (Year-over-Year) COMPARISON</th></tr>
                                            <tr>
                                                <th tkey="TrafficPercent">Traffic </th>
                                                <th tkey="WritesPercent">Writes</th>
                                                <th tkey="ClosingNoPercent">Closing%</th>
                                                <th tkey="MonthlySalesForecastPercent">Monthly Sales Forecast</th>
                                                <th tkey="TrafficPercent">Traffic</th>
                                                <th tkey="WritesPercent">Writes</th>
                                                <th tkey="ClosingNoPercent">Closing%</th>
                                            </tr>
                                        </thead>
                                        <tbody></tbody>
                                    </table>
                                </div>
                                <div id="MOMReport" style="width:100%">
                                    <table id="MOMDistrict" class="table table-striped table-bordered dt-responsive nowrap data-entry" cellspacing="0" style="display:none;">
                                        <thead>
                                            <tr><th rowspan="2" tkey="DealerNumber">Clubhouse Number</th><th rowspan="2" tkey="DealerName">Clubhouse Name</th><th colspan="7"  tkey="MoMCOMPARISON">MoM COMPARISON</th><th colspan="6" tkey="YoYMTDCOMPARISON">MTD YoY COMPARISON</th></tr>
                                            <tr>
                                                <th  tkey="DealerTrafficPercent">Clubhouse Traffic </th>
                                                <th tkey="AreaAvgTrafficPercent">Area Avg. Traffic %</th>
                                                <th tkey="DealerWritesPercent">Clubhouse Writes</th>
                                                <th tkey="AreaAvgWritesPercent">Area Avg. Writes</th>
                                                <th tkey="DealerClosingNoPercent">Clubhouse Closing%</th>
                                                <th tkey="AreaAvgClosingNoPercent">Area Avg. Closing%</th>
                                                <th tkey="MonthlySalesForecastPercent">Monthly Sales Forecast</th>
                                                <th tkey="DealerTrafficPercent">Clubhouse Traffic</th>
                                                <th  tkey="AreaAvgTrafficPercent">Area Avg. Traffic %</th>
                                                <th tkey="DealerWritesPercent">Clubhouse Writes</th>
                                                <th tkey="AreaAvgWritesPercent">Area Avg. Writes</th>
                                                <th tkey="DealerClosingNoPercent">Clubhouse Closing%</th>
                                                <th tkey="AreaAvgClosingNoPercent">Area Avg. Closing%</th>
                                            </tr>
                                        </thead>
                                        <tbody></tbody>
                                    </table>
                                    <table id="MOMRegionDistrict" class="table table-striped table-bordered dt-responsive nowrap data-entry" cellspacing="0" style="display:none;">
                                        <thead>
                                            <tr><th rowspan="2" tkey="DealerDistrict"></th><th colspan="7" tkey="MoMCOMPARISON">MOM COMPARISON</th><th colspan="6" tkey="YoYMTDCOMPARISON">YOY MTD COMPARISON</th></tr>
                                            <tr>
                                                <th tkey="TrafficPercent">Traffic </th>
                                                <th tkey="NationalAvgTrafficPercent">National Avg. Traffic</th>
                                                <th tkey="DealerWritesPercent">Writes</th>
                                                <th tkey="NationalAvgWritesPercent">National Avg. Writes</th>
                                                <th tkey="ClosingNoPercent">Closing%</th>
                                                <th tkey="NationalAvgClosingNoPercent">National Avg. Closing%</th>
                                                <th tkey="MonthlySalesForecastPercent">Monthly Sales Forecast</th>
                                                <th tkey="TrafficPercent">Traffic</th>
                                                <th tkey="NationalAvgTrafficPercent">National Avg. Traffic</th>
                                                <th tkey="WritesPercent">Writes</th>
                                                <th tkey="NationalAvgWritesPercent">National Avg. Writes</th>
                                                <th tkey="ClosingNoPercent">Closing%</th>
                                                <th tkey="NationalAvgClosingNoPercent">National Avg. Closing%</th>
                                            </tr>
                                        </thead>
                                        <tbody></tbody>
                                    </table>
                                    <table id="MOMNationalArea" class="table table-striped table-bordered dt-responsive nowrap data-entry" cellspacing="0" style="display:none;">
                                        <thead>
                                            <tr><th rowspan="2" tkey="DealerArea"></th><th colspan="4" tkey="MoMCOMPARISON">MoM COMPARISON</th><th colspan="3" tkey="YoYMTDCOMPARISON">YoY MTD COMPARISON</th></tr>
                                            <tr>
                                                <th tkey="TrafficPercent">Traffic </th>
                                                <th tkey="WritesPercent">Writes</th>
                                                <th tkey="ClosingNoPercent">Closing%</th>
                                                <th tkey="MonthlySalesForecastPercent">Monthly Sales Forecast</th>
                                                <th tkey="TrafficPercent">Traffic</th>
                                                <th tkey="WritesPercent">Writes</th>
                                                <th tkey="ClosingNoPercent">Closing%</th>
                                            </tr>
                                        </thead>
                                        <tbody></tbody>
                                    </table>
                                    <table id="MOMNationalRegion" class="table table-striped table-bordered dt-responsive nowrap data-entry" cellspacing="0" style="display:none;">
                                        <thead>
                                            <tr><th rowspan="2" tkey="DealerRegion"></th><th colspan="4" tkey="MoMCOMPARISON">MoM COMPARISON</th><th colspan="3" tkey="YoYMTDCOMPARISON">YoY MTD COMPARISON</th></tr>
                                            <tr>
                                                <th tkey="TrafficPercent">Traffic </th>
                                                <th tkey="WritesPercent">Writes</th>
                                                <th tkey="ClosingNoPercent">Closing%</th>
                                                <th tkey="MonthlySalesForecastPercent">Monthly Sales Forecast</th>
                                                <th tkey="TrafficPercent">Traffic</th>
                                                <th tkey="WritesPercent">Writes</th>
                                                <th tkey="ClosingNoPercent">Closing%</th>
                                            </tr>
                                        </thead>
                                        <tbody></tbody>
                                    </table>
                                </div>
                                
                            </div>
                        </div>
                        
                    </div>
        
            </div>
            
            
        </div>
    </div>

</body>
<style type="text/css">
    .ms-webpart-titleText.ms-webpart-titleText, .ms-webpart-titleText > a {
        background-color: white;
        /*font-family: helvetica, verdana, arial, sans-serif, Geneva, sans-serif;*/
        font-size: 22px;
        font-weight: bold;
        color: #101010 !important;
        padding: 5px 15px;
        /*border: 2px solid #008AD2;*/
        box-shadow: none;
        margin-bottom: 8px;
        /*max-width: min-content;*/
        text-align: center !important;
        line-height:40px;
        /*border-bottom: 1px solid #e3e9ee;*/ 
    }

    .ms-WPBorderBorderOnly{

       border: 1px solid #fff;
    }
</style>
`
    this.domElement.innerHTML = tabsHTML;

    // Initialize jQuery tabs
    (jQuery("#tabs", this.domElement) as any).tabs();
  }

  protected onInit(): Promise<void> {
    SPComponentLoader.loadCss('https://mazdausa.sharepoint.com/sites/MCIOneMazdaDev/trafficlog/SiteAssets/css/jsgrid.min.css');
    SPComponentLoader.loadCss('https://mazdausa.sharepoint.com/sites/MCIOneMazdaDev/trafficlog/SiteAssets/css/jsgrid-theme.min.css');
    SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/4.0.0-alpha.6/css/bootstrap.min.css');
    SPComponentLoader.loadCss('https://code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css');
    SPComponentLoader.loadCss('https://mazdausa.sharepoint.com/sites/MCIOneMazdaDev/trafficlog/SiteAssets/css/responsive.bootstrap.min.css');
    SPComponentLoader.loadCss('https://mazdausa.sharepoint.com/sites/MCIOneMazdaDev/trafficlog/SiteAssets/css/bootstrap.css');
    SPComponentLoader.loadCss('https://mazdausa.sharepoint.com/sites/MCIOneMazdaDev/trafficlog/SiteAssets/css/Site.css');
    
    SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCIOneMazdaDev/trafficlog/SiteAssets/js/jquery-1.10.2.js')
        .then(() => {
            return SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCIOneMazdaDev/trafficlog/SiteAssets/js/jsgrid.min.js');
        })
        .then(() => {
            return SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCIOneMazdaDev/trafficlog/SiteAssets/js/jquery-ui.js');

        })
        .then(() => {
          return SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCIOneMazdaDev/trafficlog/SiteAssets/js/lang.js');

      })
        .then(() => {
        return SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCIOneMazdaDev/trafficlog/SiteAssets/js/DetailReport.js');
        
    })
    .then(() => {
        return SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCIOneMazdaDev/trafficlog/SiteAssets/js/DropDown.js');
        
    })
        .then(() => {
         return SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCIOneMazdaDev/trafficlog/SiteAssets/js/SummaryReport_new.js');
        })
        .then(() => {
            return SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCIOneMazdaDev/trafficlog/SiteAssets/js/csvExport.js');
           })
  

    
    
    
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
