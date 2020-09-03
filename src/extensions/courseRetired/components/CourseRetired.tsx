import { Log } from '@microsoft/sp-core-library';
import { override } from '@microsoft/decorators';
import * as React from 'react';
import { Toggle } from "office-ui-fabric-react";
import styles from './CourseRetired.module.scss';

export interface ICourseRetiredProps {
  retired: boolean;
  onChanged(value: boolean) : void;
}

const LOG_SOURCE: string = 'CourseRetired';

export default class CourseRetired extends React.Component<ICourseRetiredProps, any> {
  constructor(props: ICourseRetiredProps){
    super(props);

    this.state= {
      checked: this.props.retired
    };
  }

  @override
  public render(): React.ReactElement<{}> {
    return (
      <div className={styles.cell}>
        <Toggle onText="Yes" offText="No" defaultChecked={ this.state.checked } 
          onChange={ (event,checked) => {
            this.setState({
              checked: !this.state.checked
            });

            if(this.props.onChanged) {
              this.props.onChanged(checked);
            }
        }} />
      </div>
    );
  }
}