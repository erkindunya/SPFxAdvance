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

export default class CourseRetired extends React.Component<ICourseRetiredProps, {}> {
  @override
  public componentDidMount(): void {
    Log.info(LOG_SOURCE, 'React Element: CourseRetired mounted');
  }

  @override
  public componentWillUnmount(): void {
    Log.info(LOG_SOURCE, 'React Element: CourseRetired unmounted');
  }

  @override
  public render(): React.ReactElement<{}> {
    return (
      <div className={styles.cell}>
        <Toggle onText="Yes" offText="No" checked={ this.props.retired } 
          onChange={ (event,checked) => {
            this.props.onChanged(checked);
        }} />
      </div>
    );
  }
}