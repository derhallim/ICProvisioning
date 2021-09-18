import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
} from '@microsoft/sp-property-pane';

import ZoneContentDisplay from './components/ZoneContentDisplay';
import { IZoneContentDisplayProps } from './components/IZoneContentDisplayProps';

export default class ZoneContentDisplayWebPart extends BaseClientSideWebPart<IZoneContentDisplayProps> {
  private listDropDownOptions: IPropertyPaneDropdownOption[] = [];
  private zoneDropDownOptions: IPropertyPaneDropdownOption[] = [];
  private zoneOptions = [];
  
  private GetListData(): any {
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
           result.value.map((items) => { if(items.BaseTemplate == "100"){ listTitle.push(items.Title); } });
          self.LoadListDropDownValues(listTitle);
        },
        (error) => {
          console.log(error);
        }
      );
  }

  private GetZonedetails(): any {
    var self = this, pages = [];
    var url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.selectedList}')/items?$select=Title`;

    fetch(url, {
      credentials: 'same-origin',
      headers: { 'Content-Type': 'application/json', 'Accept': 'application/json' }
    })
      .then((res) => res.json())
      .then(
        (result) => {
         if(result.value != undefined){
          result.value.map(item =>{
            pages.push(item.Title);
          });
          this.zoneOptions = pages.filter((item, i, ar) => ar.indexOf(item) === i);
         self.LoadZoneDropDownValues(pages.filter((item, i, ar) => ar.indexOf(item) === i));
         } 
         else{
         self.LoadZoneDropDownValues([]);
         }
        },
        (error) => {
          console.log(error);
        }
      );
  }

  private LoadListDropDownValues(data, ): any {
    this.listDropDownOptions = [];
    this.listDropDownOptions.push({ key: "Select", text: "Select" });

    data.map((items) => {
      this.listDropDownOptions.push({ key: items, text: items });
    });
}

private LoadZoneDropDownValues(data): Promise<IPropertyPaneDropdownOption[]>  {
  let values = [];
  values.push({ key: "Select", text: "Select" });
  return new Promise<IPropertyPaneDropdownOption[]>((resolve: (options: IPropertyPaneDropdownOption[]) => void, reject: (error: any) => void) => {
      data.map((items) => {
        values.push({ key: items, text: items });
      });
      this.zoneDropDownOptions = values;
      this.render();
      this.context.propertyPane.refresh();
      resolve(values);
  });
}

  public render(): void {
    const element: React.ReactElement<IZoneContentDisplayProps > = React.createElement(
      ZoneContentDisplay,
      {
        selectedList: this.properties.selectedList,
        selectedZone: this.properties.selectedZone,
        siteAbsoluteURL: this.context.pageContext.web.absoluteUrl
      }
    );

    if(this.renderedOnce === false){
      this.GetListData();
      this.GetZonedetails();
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

    if (propertyPath == 'selectedList') {
      if (newValue != oldValue) {
        this.properties.selectedList = newValue;
        this.GetZonedetails();

        super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
        const previousItem = this.properties.selectedZone;
        this.properties.selectedZone = undefined;

        this.onPropertyPaneFieldChanged('selectedZone', previousItem, this.properties.selectedZone);
        this.context.propertyPane.refresh();

        this.LoadZoneDropDownValues(this.zoneOptions)
        .then((result: IPropertyPaneDropdownOption[]): void => {
          this.zoneDropDownOptions = result;
          this.render();
          this.context.propertyPane.refresh();
        });
      
        this.context.propertyPane.refresh();

      }
    }
    else if (propertyPath == 'selectedZone') {
      if (newValue != oldValue) {
        this.properties.selectedZone = newValue;

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
                PropertyPaneDropdown('selectedList', {
                  label: "Select the list to display.",
                  options: this.listDropDownOptions,
                  selectedKey: "Select"
                }),
                PropertyPaneDropdown('selectedZone', {
                  label: "Select the zone for the content to display.",
                  options: this.zoneDropDownOptions,
                  selectedKey: "Select"
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
