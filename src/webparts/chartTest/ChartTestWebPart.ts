import * as React from 'react';
import * as ReactDom from 'react-dom';

import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import * as strings from 'ChartTestWebPartStrings';
import ChartTest from './components/ChartTest';
import { IChartTestProps, IChartData } from './components/IChartTestProps';
import { IWorkStatusItem } from '../../models';

export interface IChartTestWebPartProps {
  description: string;
}

export default class ChartTestWebPart extends BaseClientSideWebPart<IChartTestWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  public render(): void {
    
    this._getListItems()
      .then(response => {
      
        const data: IChartData = {
          labels: [],
          datasets: []
        };
        
        data.datasets.push({			
          label: "Porcentaje de realización de los Items de trabajo",
          data: response.map(item => item.PercentComplete),
          backgroundColor: response.map(item => `#${Math.floor(Math.random()*16777215).toString(16)}`),
        }); //'rgba(255, 99, 132, 0.5)'
		    data.labels = response.map(item => item.Title);

        const element: React.ReactElement<IChartTestProps> = React.createElement(
          ChartTest,
          {
            chartData: data,
            onAddListItem: this._onAddListItem,
            isDarkTheme: this._isDarkTheme,
            environmentMessage: this._environmentMessage,
            hasTeamsContext: !!this.context.sdks.microsoftTeams,
            userDisplayName: this.context.pageContext.user.displayName
          }
        );

        ReactDom.render(element, this.domElement);
      });

  }

  private _onAddListItem = (item, percentComplete): void => {
    this._addListItem(item, percentComplete)
      .then(() => {
        this._getListItems()
          .then(response => {
            this.render();
          });
      });
  }

  private _getListItems(): Promise<IWorkStatusItem[]> { 
    const endpoint: string = this.context.pageContext.web.absoluteUrl
    + `/_api/web/lists/getbytitle('Work Status')/items?$select=Id,Title, PercentComplete`;
    
    return this.context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1)
      .then(response => {
        return response.json();
      })
      .then(jsonResponse => {                
        return jsonResponse.value;
      }) as Promise<IWorkStatusItem[]>;
  }

  private _getItemEntityType(): Promise<string> {
    return this.context.spHttpClient.get(
        this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getbytitle('Work Status')?$select=ListItemEntityTypeFullName`,
        SPHttpClient.configurations.v1)
      .then(response => {
        return response.json();
      })
      .then(jsonResponse => {
        return jsonResponse.ListItemEntityTypeFullName;
      }) as Promise<string>;
  }
  
  private _addListItem(item: string, percentComplete: number): Promise<SPHttpClientResponse> {
    return this._getItemEntityType()
      .then(spEntityType => {
        const request: any = {};
        request.body = JSON.stringify({
          Title: item,
          PercentComplete: percentComplete, 
          '@odata.type': spEntityType
        });
  
        return this.context.spHttpClient.post(
          this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getbytitle('Work Status')/items`,
          SPHttpClient.configurations.v1,
          request);
        }
      ) ;
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
    this.domElement.style.setProperty('--link', semanticColors.link);
    this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered);

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
