import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'FirstWebpartWebPartStrings';
import FirstWebpart from './components/FirstWebpart';
import { IFirstWebpartProps } from './components/IFirstWebpartProps';

import { MSGraphClient } from '@microsoft/sp-http';

export interface IFirstWebpartWebPartProps {
  msGraphClient: MSGraphClient;
}

export default class FirstWebpartWebPart extends BaseClientSideWebPart <IFirstWebpartWebPartProps> {

  public render(): void {
    this.context.msGraphClientFactory.getClient().then((msGraphClient: MSGraphClient) => {
      const element: React.ReactElement<IFirstWebpartProps> = React.createElement(
        FirstWebpart,
        {
          msGraphClient
        }
      );

      ReactDom.render(element, this.domElement);
    });
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
