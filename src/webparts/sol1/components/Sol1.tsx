import * as React from 'react';
import styles from './Sol1.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';

export interface ISol1Props {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
}

export default class Sol1 extends React.Component<ISol1Props, {}> {

  private async sendData(): Promise<void> {
    const nom = (document.getElementById('name') as HTMLInputElement).value;
    const mail = (document.getElementById('email') as HTMLInputElement).value;
    const age = (document.getElementById('age') as HTMLInputElement).value;

    const url = `${this.props.context.pageContext.web.absoluteUrl}/sites/ABC/_api/web/lists/getbytitle('personne')/items`;

    const itemBody = {
      'Title': nom,
      'Email': mail,
      'Age': age
    };

    //POST REQUEST

    try {
      const postResponse = await this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
        headers: {
          'Accept': 'application/json',
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(itemBody),
      });

      if (postResponse.ok) {
        const postResponseJSON = await postResponse.json();
        console.log('POST Response:', postResponseJSON);
        alert('POST Success');
      } else {
        throw new Error(`POST Error: ${postResponse.statusText}`);
      }
    } catch (postError) {
      // console.error('POST Error:', postError);
    }
//GET REQUEST
    try {
      const getResponse = await this.props.context.spHttpClient.get(url, SPHttpClient.configurations.v1);

      if (getResponse.ok) {
        const getResponseJSON = await getResponse.json();
        // console.log('GET Response:', getResponseJSON);

        if (getResponseJSON != null && getResponseJSON.value != null) {
          getResponseJSON.value.forEach((item: any) => {
            console.log(item);
          });
        }
      } else {
        throw new Error(`GET Error: ${getResponse.statusText}`);
      }
    } catch (getError) {
      console.error('GET Error:', getError);
    }
  }

  public render(): React.ReactElement<ISol1Props> {
    const {
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.sol1} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>Web part property value: <strong>{escape(this.props.description)}</strong></div>
        </div>
        <div>
          <h3>Welcome to SharePoint Framework!</h3>
          <p>
            The SharePoint Framework (SPFx) is an extensibility model for Microsoft Viva, Microsoft Teams, and SharePoint. It's the easiest way to extend Microsoft 365 with automatic Single Sign-On, automatic hosting, and industry-standard tooling.
          </p>
          <h4>Learn more about SPFx development:</h4>
          <ul className={styles.links}>
            <li><a href="https://aka.ms/spfx" target="_blank" rel="noreferrer">SharePoint Framework Overview</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank" rel="noreferrer">Use Microsoft Graph in your solution</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank" rel="noreferrer">Build for Microsoft Teams using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank" rel="noreferrer">Build for Microsoft Viva Connections using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-store" target="_blank" rel="noreferrer">Publish SharePoint Framework applications to the marketplace</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-api" target="_blank" rel="noreferrer">SharePoint Framework API reference</a></li>
            <li><a href="https://aka.ms/m365pnp" target="_blank" rel="noreferrer">Microsoft 365 Developer Community</a></li>
          </ul>
        </div>
        <table>
          <tr>
            <td><b>Name</b></td>
            <td><input type='text' id='name' required /></td>
          </tr>
          <tr>
            <td><b>Email</b></td>
            <td><input type='email' id='email' required /></td>
          </tr>
          <tr>
            <td><b>Age</b></td>
            <td><input type='number' id='age' required /></td>
          </tr>
        </table>
        <button onClick={() => this.sendData()}>Send</button>
      </section>
    );
  }
}
