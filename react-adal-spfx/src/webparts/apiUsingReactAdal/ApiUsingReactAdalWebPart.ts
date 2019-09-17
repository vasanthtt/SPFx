import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'ApiUsingReactAdalWebPartStrings';
import ApiUsingReactAdal from './components/ApiUsingReactAdal';
import { IApiUsingReactAdalProps } from './components/IApiUsingReactAdalProps';
import { runWithAdal } from 'react-adal';
import { authContext } from '../../common/adalConfig';

export interface IApiUsingReactAdalWebPartProps {
  description: string;
}
const DO_NOT_LOGIN = false;

export default class ApiUsingReactAdalWebPart extends BaseClientSideWebPart<IApiUsingReactAdalWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IApiUsingReactAdalProps> = React.createElement(
      ApiUsingReactAdal,
      {
        description: this.properties.description
      }
    );

    runWithAdal(authContext, () => {
      ReactDom.render(element, this.domElement);
    }, DO_NOT_LOGIN);
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
