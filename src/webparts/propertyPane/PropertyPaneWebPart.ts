import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
} from '@microsoft/sp-property-pane';

import { PropertyPaneCounter } from "../../Controls/PropertyPaneCounter";

import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './PropertyPaneWebPart.module.scss';
import * as strings from 'PropertyPaneWebPartStrings';

export interface IPropertyPaneWebPartProps {
  count: number;
}

export default class PropertyPaneWebPart extends BaseClientSideWebPart<IPropertyPaneWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.propertyPane }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Property Pane Demo!</span>
              <p class="${ styles.subTitle }">Custom PropertyPane Controls.</p>
              <p class="${ styles.description }">${this.properties.count}</p>
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
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneCounter('count', {
                  label: 'Count:',
                  initialValue: this.properties.count,
                  onPropertyChanged: (newValue: number) => {
                    this.properties.count = newValue;
                    this.render();
                  }
                })
              ]
            }
          ]
        }
      ]
    };
  }
}