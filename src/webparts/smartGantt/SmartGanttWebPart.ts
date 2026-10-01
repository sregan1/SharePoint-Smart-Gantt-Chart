import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneField,
  PropertyPaneDropdown,
  PropertyPaneTextField,
  PropertyPaneToggle,
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
import { IWorkingCalendar, DEFAULT_WORKING_CALENDAR, ViewMode, ZoomLevel } from './models';
import * as strings from 'SmartGanttWebPartStrings';

export interface ISmartGanttWebPartProps {
  title: string;
  defaultView?: ViewMode;
  defaultZoom?: ZoomLevel;
  /** One toggle per weekday, 0 = Sunday … 6 = Saturday; undefined = default (Mon-Fri). */
  workDay0?: boolean;
  workDay1?: boolean;
  workDay2?: boolean;
  workDay3?: boolean;
  workDay4?: boolean;
  workDay5?: boolean;
  workDay6?: boolean;
  /** Non-working dates, one per line (yyyy-MM-dd; parsed leniently). */
  holidays?: string;
}

// Lenient holiday parsing: any separator, y-M-d with - / or ., 1-2 digit month/day.
function parseHolidays(raw: string | undefined): string[] {
  const out: string[] = [];
  (raw || '').split(/[\r\n,;]+/).forEach(line => {
    const m = /(\d{4})\s*[-/.]\s*(\d{1,2})\s*[-/.]\s*(\d{1,2})/.exec(line);
    if (!m) return;
    const y = +m[1], mo = +m[2], d = +m[3];
    const probe = new Date(Date.UTC(y, mo - 1, d));
    if (probe.getUTCFullYear() !== y || probe.getUTCMonth() !== mo - 1 || probe.getUTCDate() !== d) return;
    const iso = `${y}-${String(mo).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
    if (out.indexOf(iso) === -1) out.push(iso);
  });
  return out.sort();
}

export default class SmartGanttWebPart extends BaseClientSideWebPart<ISmartGanttWebPartProps> {
  private spService: SharePointService;
  private calendarCache: { key: string; value: IWorkingCalendar } | undefined;

  public async onInit(): Promise<void> {
    await super.onInit();
    const sp = spfi().using(SPFx(this.context));
    this.spService = new SharePointService(sp);
  }

  // Working calendar from the property pane. The object identity is kept stable
  // while the settings are unchanged so child components don't see a new prop.
  private getWorkingCalendar(): IWorkingCalendar {
    const p = this.properties;
    const flags = [p.workDay0, p.workDay1, p.workDay2, p.workDay3, p.workDay4, p.workDay5, p.workDay6];
    const workingDays: number[] = [];
    flags.forEach((f, i) => {
      const on = f === undefined ? DEFAULT_WORKING_CALENDAR.workingDays.indexOf(i) !== -1 : f;
      if (on) workingDays.push(i);
    });
    const holidays = parseHolidays(p.holidays);
    const key = `${workingDays.join('')}|${holidays.join(',')}`;
    if (!this.calendarCache || this.calendarCache.key !== key) {
      this.calendarCache = { key, value: { workingDays, holidays } };
    }
    return this.calendarCache.value;
  }

  public render(): void {
    // SPFx can call render (e.g. via a theme change) before onInit has finished;
    // this.properties and this.context aren't safe to read until then.
    if (!this.spService) return;
    const element: React.ReactElement<ISmartGanttProps> = React.createElement(SmartGantt, {
      title: this.properties.title || strings.WebPart_DefaultTitle,
      spService: this.spService,
      context: this.context,
      defaultView: this.properties.defaultView,
      defaultZoom: this.properties.defaultZoom,
      workingCalendar: this.getWorkingCalendar(),
    });
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected onThemeChanged(): void {
    // Raised before onInit; render() guards itself, and SPFx renders once
    // onInit resolves, so there's nothing to repaint yet.
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private getCalendarFields(): IPropertyPaneField<unknown>[] {
    const dayNames = [
      strings.Svc_PP_DaySun, strings.Svc_PP_DayMon, strings.Svc_PP_DayTue, strings.Svc_PP_DayWed,
      strings.Svc_PP_DayThu, strings.Svc_PP_DayFri, strings.Svc_PP_DaySat,
    ];
    const cal = this.getWorkingCalendar();
    const fields: IPropertyPaneField<unknown>[] = [];
    // Monday first, as in most working calendars.
    [1, 2, 3, 4, 5, 6, 0].forEach(d => {
      fields.push(PropertyPaneToggle(`workDay${d}`, {
        label: dayNames[d],
        checked: cal.workingDays.indexOf(d) !== -1,
      }));
    });
    fields.push(PropertyPaneTextField('holidays', {
      label: strings.Svc_PP_Holidays,
      description: strings.Svc_PP_HolidaysHint,
      multiline: true,
      rows: 5,
      value: this.properties.holidays,
    }));
    return fields;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPane_HeaderDescription },
          groups: [
            {
              groupName: strings.PropertyPane_GeneralGroupName,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.PropertyPane_TitleFieldLabel,
                  value: this.properties.title,
                }),
              ],
            },
            {
              groupName: strings.Svc_PP_DisplayGroup,
              groupFields: [
                PropertyPaneDropdown('defaultView', {
                  label: strings.Svc_PP_DefaultView,
                  selectedKey: this.properties.defaultView || 'gantt',
                  options: [
                    { key: 'gantt', text: strings.Svc_PP_ViewGantt },
                    { key: 'list', text: strings.Svc_PP_ViewList },
                    { key: 'kanban', text: strings.Svc_PP_ViewKanban },
                    { key: 'dashboard', text: strings.Svc_PP_ViewDashboard },
                    { key: 'portfolio', text: strings.Svc_PP_ViewPortfolio },
                  ],
                }),
                PropertyPaneDropdown('defaultZoom', {
                  label: strings.Svc_PP_DefaultZoom,
                  selectedKey: this.properties.defaultZoom || 'week',
                  options: [
                    { key: 'day', text: strings.Svc_PP_ZoomDay },
                    { key: 'week', text: strings.Svc_PP_ZoomWeek },
                    { key: 'month', text: strings.Svc_PP_ZoomMonth },
                    { key: 'quarter', text: strings.Svc_PP_ZoomQuarter },
                  ],
                }),
              ],
            },
            {
              groupName: strings.Svc_PP_CalendarGroup,
              groupFields: this.getCalendarFields(),
            },
          ],
        },
      ],
    };
  }
}
