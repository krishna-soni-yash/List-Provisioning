import * as React from 'react';
import styles from './ListCreation.module.scss';
import type { IListCreationProps } from './IListCreationProps';

export default class ListCreation extends React.Component<IListCreationProps> {
  public render(): React.ReactElement<IListCreationProps> {

    return (
      <section className={`${styles.listCreation}'}`}>
        <div>
          <h3>Required List Provision Completed</h3>
        </div>
      </section>
    );
  }
}
