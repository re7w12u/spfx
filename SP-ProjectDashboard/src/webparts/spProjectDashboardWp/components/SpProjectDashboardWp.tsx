import * as React from 'react';
import styles from './SpProjectDashboardWp.module.scss';
import type { ISpProjectDashboardWpProps } from './ISpProjectDashboardWpProps';
//import { escape } from '@microsoft/sp-lodash-subset';
import { Projects } from './projects/Projects';

export default class SpProjectDashboardWp extends React.Component<ISpProjectDashboardWpProps, {}> {
  public render(): React.ReactElement<ISpProjectDashboardWpProps> {
    const {
     // description,
     // isDarkTheme,
     // environmentMessage,
      hasTeamsContext,
      //userDisplayName
    } = this.props;

    return (
      <section className={`${styles.spProjectDashboardWp} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <Projects context={this.props.context}></Projects>       
        </div>
      </section>
    );
  }
}
