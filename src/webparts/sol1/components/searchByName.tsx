import * as React from 'react';
import styles from './Sol1.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface SearchProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  redirectTo: () => void;
}

export default class List extends React.Component<SearchProps> {
  render(): React.ReactElement<SearchProps> {
    const { hasTeamsContext } = this.props;

    return (
      <section className={`${styles.sol1} ${hasTeamsContext ? styles.teams : ''}`}>
        <div>
          <label htmlFor="nom">Nom:</label>
          <input type="text" id="nom" name="nom" />
        </div>
      </section>
    );
  }
}
