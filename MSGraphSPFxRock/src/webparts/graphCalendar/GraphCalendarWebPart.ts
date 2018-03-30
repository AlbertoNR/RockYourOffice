import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'GraphCalendarWebPartStrings';
import GraphCalendar from './components/GraphCalendar';
import { IGraphCalendarProps } from './components/IGraphCalendarProps';

export interface IGraphCalendarWebPartProps {
  description: string;
}

export default class GraphCalendarWebPart extends BaseClientSideWebPart<IGraphCalendarWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IGraphCalendarProps > = React.createElement(
      GraphCalendar,
      {
        description: this.properties.description,
        spContext: this.context
      }
    );

    ReactDom.render(element, this.domElement);
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
