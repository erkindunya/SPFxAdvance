import { Log } from '@microsoft/sp-core-library';
import { override } from '@microsoft/decorators';
import {
  BaseFieldCustomizer,
  IFieldCustomizerCellEventParameters
} from '@microsoft/sp-listview-extensibility';

import * as strings from 'ProductsFieldCustomizerStrings';
import styles from './ProductsFieldCustomizer.module.scss';

/**
 * If your field customizer uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IProductsFieldCustomizerProperties {
  // This is an example; replace with your own property
  sampleText?: string;
}

const LOG_SOURCE: string = 'ProductsFieldCustomizer';

export default class ProductsFieldCustomizer
  extends BaseFieldCustomizer<IProductsFieldCustomizerProperties> {

  @override
  public onInit(): Promise<void> {
    // Add your custom initialization to this method.  The framework will wait
    // for the returned promise to resolve before firing any BaseFieldCustomizer events.
    Log.info(LOG_SOURCE, 'Activated ProductsFieldCustomizer with properties:');
    Log.info(LOG_SOURCE, JSON.stringify(this.properties, undefined, 2));
    Log.info(LOG_SOURCE, `The following string should be equal: "ProductsFieldCustomizer" and "${strings.Title}"`);
    return Promise.resolve();
  }

  @override
  public onRenderCell(event: IFieldCustomizerCellEventParameters): void {
    // Color Code Price
    let tax : number = event.fieldValue as number;
    let style ="";

    if(tax < 5) {
      //event.domElement.classList.add(styles.lowvalue);
      style = "background-color: lightgreen";
    } else if(tax >= 5 && tax <15) {
      //event.domElement.classList.add(styles.medvalue);
      style = "background-color: lightskyblue";
    } else {
      //event.domElement.classList.add(styles.highvalue);
      style = "background-color: lightsalmon";
    }

    event.domElement.innerHTML = `<div style="${ style }">${ tax }</div>`;

    event.domElement.classList.add(styles.cell);
  }

  @override
  public onDisposeCell(event: IFieldCustomizerCellEventParameters): void {
    // This method should be used to free any resources that were allocated during rendering.
    // For example, if your onRenderCell() called ReactDOM.render(), then you should
    // call ReactDOM.unmountComponentAtNode() here.
    super.onDisposeCell(event);
  }
}