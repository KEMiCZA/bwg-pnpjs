import * as React from 'react';
import styles from './HelloBiwug.module.scss';
import { IHelloBiwugProps } from './IHelloBiwugProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class HelloBiwug extends React.Component<IHelloBiwugProps, {}> {
  public render(): React.ReactElement<IHelloBiwugProps> {
    return (
      <div className={ styles.helloBiwug }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Hello BIWUG 2019!</span>
              <p className={ styles.description }>{escape(this.props.description)}</p>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
