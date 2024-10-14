import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'ModernPageCopyWebPartStrings';
import ModernPageCopy from './components/ModernPageCopy';
import { IModernPageCopyProps } from './components/IModernPageCopyProps';
import { IModernPageService, ModernPageService } from './ModernPageService';

export interface IModernPageCopyWebPartProps {
  templateName: string;
  templateSiteRelativeUrl: string
  fieldTitle: string
}

export default class ModernPageCopyWebPart extends BaseClientSideWebPart<IModernPageCopyWebPartProps> {


  private _ModernPageService: IModernPageService;

  public render(): void {
    const element: React.ReactElement<IModernPageCopyProps> = React.createElement(
      ModernPageCopy,
      {
        copyPage: async (pageName: string) => {
          if (pageName !== null && pageName.length > 3) {
            const pageNameWithoutSpaces = pageName.replace(/\s/g, "")
            await this._ModernPageService.copyPage(
              this.properties.templateSiteRelativeUrl,
              this.properties.templateName,
              pageNameWithoutSpaces)

            location.href = `${this.context.pageContext.web.absoluteUrl}/SitePages/${pageNameWithoutSpaces}.aspx`;
          }
        },
        fieldTitle: this.properties.fieldTitle,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {    
    this._ModernPageService = this.context.serviceScope.consume(ModernPageService.serviceKey);
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
                PropertyPaneTextField('Title', {
                  label: "fieldTitle"
                }),
                PropertyPaneTextField('templateSiteRelativeUrl', {
                  label: "templateSiteRelativeUrl"
                }),
                PropertyPaneTextField('templateName', {
                  label: "templateName"
                })

              ]
            }
          ]
        }
      ]
    };
  }
}
