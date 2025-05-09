import * as React from 'react';
import styles from './ConnectionTester.module.scss';
import type { IConnectionTesterProps } from './IConnectionTesterProps';
import { ConnectionTest } from './ConnectionTest'; // Изменено с 'import ConnectionTest from' на 'import { ConnectionTest } from'

export default class ConnectionTester extends React.Component<IConnectionTesterProps, {}> {
  public render(): React.ReactElement<IConnectionTesterProps> {
    const {
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
      context
    } = this.props;

    return (
      <section className={`${styles.connectionTester} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <h2>Connection Tester Tool</h2>
          <p>This tool helps you test the connection to a SharePoint site and check lists.</p>
        </div>
        
        <ConnectionTest context={context} />
        
        <div style={{ marginTop: '20px', borderTop: '1px solid #ccc', paddingTop: '10px' }}>
          <p>
            Logged in as: <strong>{userDisplayName}</strong>
          </p>
          <p>
            Environment: <strong>{environmentMessage}</strong>
          </p>
        </div>
      </section>
    );
  }
}