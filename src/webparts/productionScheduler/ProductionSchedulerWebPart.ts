import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'ProductionSchedulerWebPartStrings';
import { SchedulerEngine } from './SchedulerEngine';
import { SharePointDataService } from './SharePointDataService';

export interface IProductionSchedulerWebPartProps {
  description: string;
}

export default class ProductionSchedulerWebPart extends BaseClientSideWebPart<IProductionSchedulerWebPartProps> {
  private schedulerEngine: SchedulerEngine | undefined;
  private initialized: boolean = false;

  public render(): void {
    if (this.initialized) {
      return;
    }

    this.domElement.innerHTML = `<div id="scheduler-root"></div>`;

    const root = this.domElement.querySelector('#scheduler-root') as HTMLElement;

    const dataService = new SharePointDataService(
      this.context.pageContext.web.absoluteUrl,
      this.context.spHttpClient,
      this.properties.description || 'Project Name'
    );

    this.schedulerEngine = new SchedulerEngine(root, dataService);
    this.schedulerEngine.init().catch((error: Error): void => {
      this.domElement.innerHTML = `<div>Failed to initialize scheduler: ${error.message}</div>`;
    });

    this.initialized = true;
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