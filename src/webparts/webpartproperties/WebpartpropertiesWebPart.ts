import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneToggle
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './WebpartpropertiesWebPart.module.scss';
import * as strings from 'WebpartpropertiesWebPartStrings';

export interface IWebpartpropertiesWebPartProps {
  wpdescription : string;
  wptitle : string;
  wpsubtitle : string;
  opentouseonanycontent : boolean;
  selectbusinessunit : string;
  arenominationsopen : boolean;
  
}

export default class WebpartpropertiesWebPart extends BaseClientSideWebPart<IWebpartpropertiesWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.webpartproperties }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">${escape(this.properties.wptitle)}</span>
              <p class="${ styles.subTitle }">${escape(this.properties.wpsubtitle)}</p>
              <p class="${ styles.description }">${escape(this.properties.wpdescription)}</p>
              <p class="${ styles.description }">Is this webpart open to use for any department : ${this.properties.opentouseonanycontent}</p>
              <p class="${ styles.description }">Please select your Business Unit : ${this.properties.selectbusinessunit}</p>
              <p class="${ styles.description }">Are the listed department accepting nominations : ${this.properties.arenominationsopen}</p>
              <a href="https://aka.ms/spfx" class="${ styles.button }">
                <span class="${ styles.label }">Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>`;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "ABC Corp Registration WP"
          },
          groups: [
            {
              groupName: "Custom Webparts",
              groupFields: [
                PropertyPaneTextField('wptitle', {
                  label: 'WebPart Title'
                }),
                PropertyPaneTextField('wpsubtitle', {
                  label: 'WebPart Sub Title'
                }),
                PropertyPaneTextField('wpdescription', {
                  label: 'WebPart Description', multiline : true
                }),
                PropertyPaneCheckbox('opentouseonanycontent',{
                  text: 'Is this webpart open to use'
                }),
                PropertyPaneDropdown('selectbusinessunit',{
                  label: 'Select an option', 
                  options : [
                    {key: 'HR', text: 'HR'},
                    {key: 'Sales', text: 'Sales'},
                    {key: 'Marketing', text: 'Marketing'},
                    {key: 'IT', text: 'IT'},
                  ]
                }),
                PropertyPaneToggle('arenominationsopen',{
                  label : 'Are Nominations open',
                  onText: 'On',
                  offText: 'Off'
                })                
              ]
            }
          ]
        }
      ]
    };
  }
}
