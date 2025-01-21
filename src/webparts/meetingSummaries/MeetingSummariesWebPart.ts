import * as React from 'react';
import * as ReactDom from 'react-dom';
import '../../globalStyles/hideSpPageStyles.scss'
import '../../globalStyles/workbench.scss'
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy, PropertyFieldNumber } from '@pnp/spfx-property-controls';
import { SPFI } from '@pnp/sp';
import * as strings from 'MeetingSummariesWebPartStrings';
import MeetingSummaries, { IMeetingSummariesProps } from './components/MeetingSummaries';
import MeetingSummariesEdit, { IMeetingSummariesEditProps } from './components/MeetingSummariesEdit';
import { getSP } from '../../pnp.config';

const { solution } = require('../../../config/package-solution.json');

export interface IMeetingSummariesWebPartProps {
  Title: string;
  MeetingSummariesListId: string
  CompaniesList: string
  ExternalUsersOptions: string
  TasksListId: string
  sp: SPFI;
  context: WebPartContext;
  FormID?: string;
}

export default class MeetingSummariesWebPart extends BaseClientSideWebPart<IMeetingSummariesWebPartProps> {

  private sp: SPFI

  public render(): void {

    const url = new URL(window.location.href);
    const FormID = url.searchParams.get("FormID");

    if (FormID === null) {
      // FormID exists, do something
      const element: React.ReactElement<IMeetingSummariesProps> = React.createElement(
        MeetingSummaries,
        {
          userDisplayName: this.context.pageContext.user.displayName,
          Title: this.properties.Title,
          MeetingSummariesListId: this.properties.MeetingSummariesListId,
          CompaniesList: this.properties.CompaniesList,
          ExternalUsersOptions: this.properties.ExternalUsersOptions,
          TasksListId: this.properties.TasksListId,
          sp: this.sp,
          context: this.context
        }
      );
      ReactDom.render(element, this.domElement);
    } else {
      // FormID is not present in the URL
      const elementEdit: React.ReactElement<IMeetingSummariesProps> = React.createElement(
        MeetingSummariesEdit,
        {
          userDisplayName: this.context.pageContext.user.displayName,
          Title: this.properties.Title,
          MeetingSummariesListId: this.properties.MeetingSummariesListId,
          CompaniesList: this.properties.CompaniesList,
          ExternalUsersOptions: this.properties.ExternalUsersOptions,
          TasksListId: this.properties.TasksListId,
          sp: this.sp,
          context: this.context,
          FormID: FormID
        }
      );
      ReactDom.render(elementEdit, this.domElement);
    }
  }

  protected onInit(): Promise<void> {
    console.log(solution.name, solution.version);
    this.sp = getSP(this.context)
    return this._getEnvironmentMessage().then(message => { });
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

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse(solution.version);
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
                PropertyPaneTextField('Title', {
                  label: "Title",
                  value: this.properties.Title
                }),
                PropertyFieldListPicker("MeetingSummariesListId", {
                  label: "MeetingSummariesListId",
                  key: "MeetingSummariesListId",
                  selectedList: this.properties.MeetingSummariesListId,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context as any,
                  deferredValidationTime: 0,
                }),
                PropertyFieldListPicker("CompaniesList", {
                  label: "CompaniesList",
                  key: "CompaniesList",
                  selectedList: this.properties.CompaniesList,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context as any,
                  deferredValidationTime: 0,
                }),
                PropertyFieldListPicker("ExternalUsersOptions", {
                  label: "ExternalUsersOptions",
                  key: "ExternalUsersOptions",
                  selectedList: this.properties.ExternalUsersOptions,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context as any,
                  deferredValidationTime: 0,
                }),
                PropertyFieldListPicker("TasksListId", {
                  label: "TasksListId",
                  key: "TasksListId",
                  selectedList: this.properties.TasksListId,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context as any,
                  deferredValidationTime: 0,
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
