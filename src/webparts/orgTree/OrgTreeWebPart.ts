import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import OrgTree from './components/OrgTree';
import { IOrgTreeProps } from './components/IOrgTreeProps';

export default class OrgTreeWebPart extends BaseClientSideWebPart<IOrgTreeProps> {

  public render(): void {
    const element: React.ReactElement<IOrgTreeProps> = React.createElement(
      OrgTree,
      {
        // ... existing props ...
        listTitle: this.properties.listTitle || "Employee Directory v2",
        
        colTitle: this.properties.colTitle || "Title",
        colJob: this.properties.colJob || "JobTitle",
        colDept: this.properties.colDept || "Department",
        colSuperior: this.properties.colSuperior || "Superior_Department",
        colEmail: this.properties.colEmail || "Email",
        colMobile: this.properties.colMobile || "MobilePhone",
        
        // NEW Mapping
        colJobRank: this.properties.colJobRank || "Job_Position_Code", // <--- ADD THIS LINE
        
        colContractType: this.properties.colContractType || "Contract_Type",
        contractTypeFilter: this.properties.contractTypeFilter || "UG1,UG2", 
        
        // ... existing props ...
        webPartWidth: this.properties.webPartWidth || 100,
        transparentBackground: this.properties.transparentBackground || false,
        context: this.context
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
          header: { description: "Configure your Organization Chart" },
          groups: [
            // ... Group 1 ...
            {
              groupName: "Column Mappings (Internal Names)",
              groupFields: [
                PropertyPaneTextField('colTitle', { label: "Employee Name Column" }),
                PropertyPaneTextField('colJob', { label: "Job Title Column" }),
                PropertyPaneTextField('colDept', { label: "Department Column" }),
                PropertyPaneTextField('colSuperior', { label: "Superior Dept Column" }),
                // NEW FIELD
                PropertyPaneTextField('colJobRank', { 
                    label: "Job Rank/Code Column",
                    description: "Used for sorting (e.g., Job_Position_Code)"
                }), 
                PropertyPaneTextField('colEmail', { label: "Person/Email Column" }),
                PropertyPaneTextField('colMobile', { label: "Mobile Column" }),
                PropertyPaneTextField('colContractType', { label: "Contract Type Column" })
              ]
            },
            {
              groupName: "Logic & Filters",
              groupFields: [
                PropertyPaneTextField('contractTypeFilter', { 
                  label: "Valid Contract Types",
                  description: "Comma-separated codes (e.g., 'UG1, UG2'). Employees NOT in this list will be moved to 'DCT Saradnici'."
                })
              ]
            },
            {
              groupName: "Look & Feel",
              groupFields: [
                PropertyPaneSlider('webPartWidth', {
                  label: "Web Part Width (%)",
                  min: 50,
                  max: 100,
                  step: 5
                }),
                PropertyPaneToggle('transparentBackground', {
                  label: "Transparent Background",
                  onText: "Blend",
                  offText: "White/Gray Box"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}