import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'PnPDemo01WebPartStrings';
import PnPDemo01 from './components/PnPDemo01';
import { IPnPDemo01Props } from './components/IPnPDemo01Props';
import { setup as pnpSetup } from "@pnp/common";

export interface IPnPDemo01WebPartProps {
  description: string;
}

export default class PnPDemo01WebPart extends BaseClientSideWebPart<IPnPDemo01WebPartProps> {

  public onInit(): Promise<void> {

    return super.onInit().then(_ => {
      pnpSetup({
        spfxContext: this.context
      });
    });
  }

  public render(): void {
    const element: React.ReactElement<IPnPDemo01Props> = React.createElement(
      PnPDemo01,
      {
        description: this.properties.description
      }
    );

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
