import * as React from 'react';
import * as ReactDOM from 'react-dom';

import { Log } from '@microsoft/sp-core-library';
import { override } from '@microsoft/decorators';
import {
  BaseFieldCustomizer,
  IFieldCustomizerCellEventParameters
} from '@microsoft/sp-listview-extensibility';
import { sp, IItemUpdateResult } from "@pnp/sp/presets/all";
import * as strings from 'CourseRetiredFieldCustomizerStrings';
import CourseRetired, { ICourseRetiredProps } from './components/CourseRetired';

/**
 * If your field customizer uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
/**
 * If your field customizer uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ICourseRetiredFieldCustomizerProperties {
  // This is an example; replace with your own property
  sampleText?: string;
}

const LOG_SOURCE: string = 'CourseRetiredFieldCustomizer';

export default class CourseRetiredFieldCustomizer
  extends BaseFieldCustomizer<ICourseRetiredFieldCustomizerProperties> {

  @override
  public onInit(): Promise<void> {
    sp.setup({
      spfxContext: this.context
    });

    // Add your custom initialization to this method.  The framework will wait
    // for the returned promise to resolve before firing any BaseFieldCustomizer events.
    Log.info(LOG_SOURCE, 'Activated CourseRetiredFieldCustomizer with properties:');
    Log.info(LOG_SOURCE, JSON.stringify(this.properties, undefined, 2));
    Log.info(LOG_SOURCE, `The following string should be equal: "CourseRetiredFieldCustomizer" and "${strings.Title}"`);
    return Promise.resolve();
  }

  @override
  public onRenderCell(event: IFieldCustomizerCellEventParameters): void {
    // Use this method to perform your custom cell rendering.
    const retired: string = event.fieldValue;

    const courseRetired: React.ReactElement<{}> =
      React.createElement(CourseRetired, 
        { 
          id: parseInt(event.listItem.getValueByName("ID").toString()),
          retired: retired=="Yes" ? true: false,
          onChanged: (checked: boolean, id: number) => {
            // Use PnP to Update
            sp.web.lists.getByTitle('Courses').items.getById(id)
              .get()
              .then((item : any) => {
                item.Retired = checked;

                return sp.web.lists.getByTitle('Courses').items.getById(id)
                          .update(item);
              })
              .then((res: IItemUpdateResult) =>{
                console.log("Course Returied: " + checked + " Status updated!");
              })
              .catch(err=> {
                console.log("Error updating retired status");
              });
          }
        } as ICourseRetiredProps);

    ReactDOM.render(courseRetired, event.domElement);
  }

  @override
  public onDisposeCell(event: IFieldCustomizerCellEventParameters): void {
    // This method should be used to free any resources that were allocated during rendering.
    // For example, if your onRenderCell() called ReactDOM.render(), then you should
    // call ReactDOM.unmountComponentAtNode() here.
    ReactDOM.unmountComponentAtNode(event.domElement);
    super.onDisposeCell(event);
  }
}