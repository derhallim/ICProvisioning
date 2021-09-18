import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
} from '@microsoft/sp-property-pane';

import FileViewer from './components/FileViewer';
import { IFileViewerProps } from './components/IFileViewerProps';

export default class FileViewerWebPart extends BaseClientSideWebPart<IFileViewerProps> {

  private libraryDropDownOptions: IPropertyPaneDropdownOption[] = [];

  private GetLibraryData(): any {
    var self = this;
    var url = this.context.pageContext.web.absoluteUrl + "/_api/web/lists?$filter=Hidden eq false";

    fetch(url, {
      credentials: 'same-origin',
      headers: { 'Content-Type': 'application/json', 'Accept': 'application/json' }
    })
      .then((res) => res.json())
      .then(
        (result) => {
          var listTitle = [];
          result.value.map((items) => { if (items.BaseTemplate == "101") { listTitle.push(items.Title); } });
          self.LoadLibraryDropDownValues(listTitle);
        },
        (error) => {
          console.log(error);
        }
      );
  }

  private LoadLibraryDropDownValues(data,): any {
    this.libraryDropDownOptions = [];
    this.libraryDropDownOptions.push({ key: "Select", text: "Select" });

    data.map((items) => {
      this.libraryDropDownOptions.push({ key: items, text: items });
    });

  }

  public render(): void {
    const element: React.ReactElement<IFileViewerProps> = React.createElement(
      FileViewer,
      {
        selectedLibrary: this.properties.selectedLibrary,
        siteAbsoluteURL: this.context.pageContext.web.absoluteUrl
      }
    );

    if (this.renderedOnce === false) {
      this.GetLibraryData();
    }

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any) {
    if (propertyPath == 'selectedLibrary') {
      if (newValue != oldValue) {
        this.properties.selectedLibrary = newValue;

        this.context.propertyPane.refresh();
      }
    }
    this.context.propertyPane.refresh();
  }
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupFields: [
                PropertyPaneDropdown('selectedLibrary', {
                  label: "Select the library to display the file.",
                  options: this.libraryDropDownOptions,
                  selectedKey: "Select"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
