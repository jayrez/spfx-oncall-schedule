import * as React from 'react';
import * as ReactDom from 'react-dom';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import OnCallSchedule from './components/OnCallSchedule';

export interface IWeeklyOnCallScheduleWebPartProps {}

export default class WeeklyOnCallScheduleWebPart extends BaseClientSideWebPart<IWeeklyOnCallScheduleWebPartProps> {

  public render(): void {
    const element = React.createElement(OnCallSchedule, {
      context: this.context
    });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }
}
