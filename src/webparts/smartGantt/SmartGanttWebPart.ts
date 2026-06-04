import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from '@microsoft/sp-property-pane';

import { spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';
import '@pnp/sp/views';

import SmartGantt from './components/SmartGantt';
import { ISmartGanttProps } from './components/SmartGantt';
import { SharePointService } from './services/SharePointService';

export interface ISmartGanttWebPartProps {
  title: string;
}

export default class SmartGanttWebPart extends BaseClientSideWebPart<ISmartGanttWebPartProps> {
  private spService: SharePointService;

  public async onInit(): Promise<void> {
    await super.onInit();
    const sp = spfi().using(SPFx(this.context));
    this.spService = new SharePointService(sp);
  }

  public render(): void {
    const element: React.ReactElement<ISmartGanttProps> = React.createElement(SmartGantt, {
      title: this.properties.title || 'Smart Gantt Chart',
      spService: this.spService,
      context: this.context,
    });
    ReactDom.render(element, this.domElement);
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
          header: { description: 'Smart Gantt Chart Settings' },
          groups: [
            {
              groupName: 'General',
              groupFields: [
                PropertyPaneTextField('title', {
                  label: 'Web Part Title',
                  value: this.properties.title,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
