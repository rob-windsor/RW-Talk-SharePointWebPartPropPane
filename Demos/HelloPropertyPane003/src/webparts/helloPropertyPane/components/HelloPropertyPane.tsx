import * as React from 'react';
import styles from './HelloPropertyPane.module.scss';
import type { IHelloPropertyPaneProps } from './IHelloPropertyPaneProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class HelloPropertyPane extends React.Component<IHelloPropertyPaneProps> {
  public render(): React.ReactElement<IHelloPropertyPaneProps> {
    const {
      description,
      color,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.helloPropertyPane} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>Description property value: <strong>{escape(description)}</strong></div>
          <div>Color property value: <strong>{escape(color)}</strong></div>
        </div>
      </section>
    );
  }
}
