import * as React from 'react';
import styles from './Sol1.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

interface IPerson {
  Title: string;
  Email: string;
  Age: number;
  ID: number;
}

export interface ISol1Props {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  redirectTo: () => void;
}

interface ISol1State {
  nom: string;
  Email: string; 
  age: string;
  ageErrorMessage: string;
}

export default class Sol1 extends React.Component<ISol1Props, ISol1State> {
  constructor(props: ISol1Props) {
    super(props);

    this.state = {
      nom: '',
      age: '',
      ageErrorMessage: '',
      Email: ''
    };
  }
  
  componentDidMount() {
    this.toUpdate();
  }

  private toUpdate(): void {
    const itemId = localStorage.getItem("ID");
    const url = `https://mch12.sharepoint.com/sites/ABC/_api/web/lists/getbytitle('personne')/items?$filter=ID eq ${itemId}`;

    // GET Api
    this.props.context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
      headers: {
        'Accept': 'application/json',
        'Content-Type': 'application/json',
      },
    })
      .then((response: SPHttpClientResponse) => {
        console.log(`Status code: ${response.status}`);
        console.log(`Status text: ${response.statusText}`);

        if (response.ok) {
          response.json().then((responseJSON: { value: IPerson[] }) => {
            console.log(responseJSON);
            // Assuming you want to update the state with the data of the first item
            if (responseJSON.value.length > 0) {
              const { Title, Age, Email } = responseJSON.value[0];
              this.setState({ nom: Title, age: Age.toString(), Email: Email });
            } else {
              console.warn('No data found for the specified ID.');
            }
          });
        } else {
          console.error('Error:', response.statusText);
        }
      })
      .catch((error) => {
        console.error('Error:', error);
      });
  }
  
  private handleNameChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    this.setState({ nom: event.target.value });
  };

  private handleAgeChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const age = event.target.value;
    this.setState({ age });

    // Vérifiez si l'âge est inférieur à 18 et mettez à jour le message d'erreur
    if (age !== '' && parseInt(age, 10) < 18) {
      this.setState({ ageErrorMessage: "You must be 18 or older to proceed." });
    } else {
      this.setState({ ageErrorMessage: '' });
    }
  };

  private handleEmailChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    this.setState({ Email: event.target.value });
  };

  // UPDATE DATA
  private async updateData(): Promise<void> {
    const { nom, age } = this.state;
    const mail = (document.getElementById('email') as HTMLInputElement).value;
  
    // Vérifiez l'âge avant de mettre à jour les données
    if (age !== '' && parseInt(age, 10) < 18) {
      this.setState({ ageErrorMessage: "You must be 18 or older to proceed." });
      return;
    }
  
    // Vérifiez le nom et le courriel
    if (nom === '' || !(/^[A-Za-z\s]+$/.test(nom))) {
      alert('le nom est vide');
      return;
    } else if (mail === '' || !(/^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/.test(mail))) {
      alert("Verify email");
      return;
    } else if (!age) {
      alert("Verify age");
      return;
    }
  
    const itemId = localStorage.getItem("ID");
  
    // Mettre à jour les données avec la méthode POST et X-HTTP-Method: MERGE
    const url = `https://mch12.sharepoint.com/sites/ABC/_api/web/lists/getbytitle('personne')/items(${itemId})`;
  
    const itemBody = {
      'Title': nom,
      'Email': mail,
      'Age': age
    };
    console.log(JSON.stringify(itemBody));
  
    try {
      const postResponse = await this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
        headers: {
          'Accept': 'application/json',
          'Content-Type': 'application/json',
          'IF-MATCH': '*',
          'X-HTTP-Method': 'MERGE'  // Use MERGE to update an existing item
        },
        body: JSON.stringify(itemBody),
      });
  
      if (postResponse.ok) {
        // La réponse JSON peut être vide, nous n'avons pas besoin de la traiter
        console.log('POST Response:', postResponse);
        this.props.redirectTo();
      } else {
        throw new Error(`POST Error: ${postResponse.statusText}`);
      }
    } catch (postError) {
      console.error('POST Error:', postError);
    }
  }

  // add data
  private async sendData(): Promise<void> {
    const { nom, age } = this.state;
    const mail = (document.getElementById('email') as HTMLInputElement).value;

    // Vérifiez l'âge avant d'envoyer les données
    if (age !== '' && parseInt(age, 10) < 18) {
      this.setState({ ageErrorMessage: "You must be 18 or older to proceed." });
      return;
    }

    // Vérifiez le nom et le courriel
    if (nom === '' || !(/^[A-Za-z\s]+$/.test(nom))) {
      alert('le nom est vide');
      return;
    } else if (mail === '' || !(/^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/.test(mail))) {
      alert("Verify email");
      return;
    } else if (!age) {
      alert("Verify age");
      return;
    }

    // Envoyer les données
    const url = `https://mch12.sharepoint.com/sites/ABC/_api/web/lists/getbytitle('personne')/items`;

    const itemBody = {
      'Title': nom,
      'Email': mail,
      'Age': age
    };

    try {
      const postResponse = await this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
        headers: {
          'Accept': 'application/json',
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(itemBody),
      });

      const postResponseText = await postResponse.text();

      if (postResponseText.trim() !== '') {
        const postResponseJSON = JSON.parse(postResponseText);
        console.log('POST Response:', postResponseJSON);
        this.props.redirectTo();
      } else {
        console.error('POST Error: Empty JSON response');
      }
    } catch (postError) {
      console.error('POST Error:', postError);
    }
  }

  public render(): React.ReactElement<ISol1Props> {
    const { hasTeamsContext, userDisplayName } = this.props;
    const { ageErrorMessage } = this.state;

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
                <td><input className="form-control" type='text' id='name' onChange={this.handleNameChange} value={this.state.nom} required /></td>
              </tr>
              <tr>
                <td><b>Email</b></td>
                <td><input className="form-control" type='email' id='email' onChange={this.handleEmailChange} value={this.state.Email} required /></td>
              </tr>
              <tr>
                <td><b>Age</b></td>
                <td>
                  <input className="form-control" type='number' id='age' onChange={this.handleAgeChange} value={this.state.age} required />
                  {ageErrorMessage && <p style={{ color: 'red' }}>{ageErrorMessage}</p>}
                </td>
              </tr>
            </tbody>
          </table>
          {localStorage.getItem("ID") ? (
            <button className="btn btn-primary" onClick={() => this.updateData()}>Update</button>
          ) : (
            <button className="btn btn-primary" onClick={() => this.sendData()}>Send</button>
          )}
        </div>
      </section>
    );
  }
}
