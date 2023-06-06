import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DisplayMode } from '@microsoft/sp-core-library';

import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'BblpopupWebPartStrings';
import Bblpopup from './components/Bblpopup';
import { IBblpopupProps } from './components/IBblpopupProps';
import { PropertyFieldDateTimePicker,DateConvention, IDateTimeFieldValue } from '@pnp/spfx-property-controls/lib/PropertyFieldDateTimePicker';
import * as moment from 'moment';
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
export interface IBblpopupWebPartProps {
  description: string;
  url:string;
  eventStartDate: IDateTimeFieldValue ;
  eventEndDate: IDateTimeFieldValue;
  webpartid: string;
}

export default class BblpopupWebPart extends BaseClientSideWebPart<IBblpopupWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _webpartid : string = '';
  private _placeholderComponent: any;
  constructor(props: IBblpopupWebPartProps) {
    super()
    
  }

  public render(): void {
   
   let renderRootElement: JSX.Element = null;
    if(this.displayMode === DisplayMode.Read){
      const bblPopupcomponent: React.ReactElement<IBblpopupProps> = React.createElement(
      Bblpopup,
      {
        url : this.properties.url,
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        eventStartDate: this.properties.eventStartDate,
        eventEndDate: this.properties.eventEndDate,
        webpartid: this._webpartid
      
      }
    );
    renderRootElement = bblPopupcomponent;
  }else if(this.displayMode === DisplayMode.Edit){
     const placeholder: React.ReactElement<any> = React.createElement(
      this._placeholderComponent,
      {
          iconName: 'BeerMug',
          iconText: 'Configure your bblpopup webpart',
          description: 'Please Config the web part',
          buttonLabel: 'Configure',
          onConfigure: () => { this.context.propertyPane.open(); }
      }
  );
    
     renderRootElement = placeholder;
  }
  else
  {
  
  renderRootElement = null;
  }
  const element = React.createElement(
    'div',
    {className: ''},
    renderRootElement
  );
    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
     
   // console.log("on init ");
    if (!this.properties.eventStartDate){
      this.properties.eventStartDate = { value: moment().toDate(), displayValue: moment().format('LL')};
    }
    if (!this.properties.eventEndDate){
      this.properties.eventEndDate = { value: moment().endOf('month').toDate(), displayValue: moment().format('LL')};
    }
    this._webpartid = this.context.instanceId.toString();
    //this.properties.webpartid = new Date().getTime.toString();

    if(this.displayMode === DisplayMode.Edit){
      this._placeholderComponent = Placeholder;
    }
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
 /**
   *
   *
   * @private
   * @param {string} date
   * @returns
   * @memberof CalendarWebPart
   */
 private onEventStartDateValidation(date:string):string{
  
  if (date && this.properties.eventEndDate.value){
    if (moment(date).isAfter(moment(this.properties.eventEndDate.value))){
      return 'invalid start date'
    }
  }
  return '';
}

/**
 *
 * @private
 * @param {string} date
 * @returns
 * @memberof CalendarWebPart
 */
private onEventEndDateValidation(date:string):string{
  if (date && this.properties.eventEndDate.value){
    if (moment(date).isBefore( moment(this.properties.eventStartDate.value))){
      return 'invalid enddate';
    }
  }
  return '';
}

protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
 // this.onEventEndDateValidation(this.properties.eventEndDate.value.toString());
  //this.onEventStartDateValidation(this.properties.eventStartDate.value.toString());
   //console.log("Field Change "+ propertyPath);
   //if(propertyPath === 'url'){
   // console.log("change pro"+ propertyPath);
   //this._webpartid = new Date().getTime().toString();
  // }
}
  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
                }),
                PropertyPaneTextField('url',{
                  label:strings.URLFieldLable              
                }),
                PropertyFieldDateTimePicker('eventStartDate', {
                  label: 'From',
                  initialDate: this.properties.eventStartDate,
                  dateConvention: DateConvention.Date,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  onGetErrorMessage: this.onEventStartDateValidation,
                  deferredValidationTime: 0,
                  key: 'eventStartDateId'
                }),
                PropertyFieldDateTimePicker('eventEndDate', {
                  label: 'to',
                  initialDate:  this.properties.eventEndDate,
                  dateConvention: DateConvention.Date,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  onGetErrorMessage:  this.onEventEndDateValidation,
                  deferredValidationTime: 0,
                  key: 'eventEndDateId'
                }),   
              ]
            }
   
          ]
        }
      ]
    };
  }
}
