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
        // alert('POST Success');
        
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
     
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.sol1} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={`container ${styles.welcome}`}>
          <h2>Hello, {escape(userDisplayName)}! You can add an item to the list</h2>
        </div>
        <div className="container">
          <table className="table">
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
          </table>
          <button className="btn btn-primary" onClick={() => this.sendData()}>Send</button>
        </div>
      </section>
    );
  }
}
