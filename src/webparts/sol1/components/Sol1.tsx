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
  redirectTo: () => void;
}

export default class Sol1 extends React.Component<ISol1Props, {}> {
  private async sendData(): Promise<void> {
    const nom = (document.getElementById('name') as HTMLInputElement).value;
    const mail = (document.getElementById('email') as HTMLInputElement).value;
    const age = (document.getElementById('age') as HTMLInputElement).value;

    //verify name
    if (nom === '' || !(/^[A-Za-z\s]+$/.test(nom))) {
      alert('le nom est vide');
      return;
    } else if (mail === '' || !(/^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/.test(mail))) {
      alert("Verify email");
      return;
    }else if (!age ) {
      alert ("Verify age");
      return;
    }


    const url = `https://mch12.sharepoint.com/sites/ABC/_api/web/lists/getbytitle('personne')/items`;

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
        this.props.redirectTo();
      } else {
        throw new Error(`POST Error: ${postResponse.statusText}`);
      }
    } catch (postError) {
      console.error('POST Error:', postError);
    }
  }

  public render(): React.ReactElement<ISol1Props> {
    const { hasTeamsContext, userDisplayName } = this.props;

    return (
      <section className={`${styles.sol1} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={`container ${styles.welcome}`}>
          <h2>Hello, {escape(userDisplayName)}! You can add an item to the list</h2>
        </div>
        <div className="container">
          <table className="table">
            <tbody>
              <tr>
                <td><b>Name</b></td>
                <td><input className="form-control" type='text' id='name' required /></td>
              </tr>
              <tr>
                <td><b>Email</b></td>
                <td><input className="form-control" type='email' id='email' required /></td>
              </tr>
              <tr>
                <td><b>Age</b></td>
                <td><input className="form-control" type='number' id='age' required /></td>
              </tr>
            </tbody>
          </table>
          <button className="btn btn-primary" onClick={() => this.sendData()}>Send</button>
        </div>
      </section>
    );
  }
}
