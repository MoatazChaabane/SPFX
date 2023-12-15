import * as React from 'react';
import styles from './Sol1.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

interface IPerson {
  Title: string;
  Email: string;
  Age: number;
  ID: number;
}

export interface IListProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  redirectTo: () => void;
}

interface IListState {
  people: IPerson[];
}

export default class List extends React.Component<IListProps, IListState> {
  constructor(props: IListProps) {
    super(props);

    this.state = {
      people: [],
    };
  }

  // DELETE DATA
  private async deleteData(itemId: number): Promise<void> {
    // Utiliser la fonction window.confirm pour afficher une boîte de dialogue de confirmation
    const confirmDelete = window.confirm('Are you sure you want to delete this item?');

    // Vérifier si l'utilisateur a confirmé la suppression
    if (!confirmDelete) {
      return;
    }

    const url = `https://mch12.sharepoint.com/sites/ABC/_api/web/lists/getbytitle('personne')/items(${itemId})`;

    try {
      const deleteResponse = await this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
        headers: {
          'Accept': 'application/json',
          'Content-Type': 'application/json',
          'IF-MATCH': '*',
          'X-HTTP-Method': 'DELETE'  // Utilisez DELETE pour supprimer un élément existant
        },
      });

      if (deleteResponse.ok) {
        // La réponse JSON peut être vide, nous n'avons pas besoin de la traiter
        console.log('DELETE Response:', deleteResponse);
        // Actualisez les données après la suppression si nécessaire
        this.getData();
      } else {
        throw new Error(`DELETE Error: ${deleteResponse.statusText}`);
      }
    } catch (deleteError) {
      console.error('DELETE Error:', deleteError);
    }
  }

  // EDIT DATA
  private async onEdit($id: number): Promise<void> {
    localStorage.setItem("ID", String($id));
    this.props.redirectTo();
  }

  // RETURN TO MAIN VIEW
  private async onReturn(): Promise<void> {
    localStorage.removeItem("ID");

    this.props.redirectTo();
  }

  componentDidMount() {
    this.getData();
  }

  private getData(): void {
    const url = `https://mch12.sharepoint.com/sites/ABC/_api/web/lists/getbytitle('personne')/items?$top=1000`;

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
            this.setState({ people: responseJSON.value });
          });
        } else {
          console.error('Error:', response.statusText);
        }
      })
      .catch((error) => {
        console.error('Error:', error);
      });
  }

  public render(): React.ReactElement<IListProps> {
    const { hasTeamsContext } = this.props;
    const { people } = this.state;

    return (
      <section className={`${styles.sol1} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <table className="table table-striped">
            <thead>
              <tr>
                <th scope="col">Title</th>
                <th scope="col">Email</th>
                <th scope="col">Age</th>
                <th scope="col">Action</th>
              </tr>
            </thead>
            <tbody>
              {people.map((person, index) => (
                <tr key={index}>
                  <td>{person.Title}</td>
                  <td>{person.Email}</td>
                  <td>{person.Age}</td>
                  <td>
                    <button onClick={() => this.onEdit(person.ID)}>
                      <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" className="bi bi-pen" viewBox="0 0 16 16">
                        <path d="M13.498.795.149-.149a1.207 1.207 0 1 1 1.707 1.708l-.149.148a1.5 1.5 0 0 1-.059 2.059L4.854 14.854a.5.5 0 0 1-.233.131l-4 1a.5.5 0 0 1-.606-.606l1-4a.5.5 0 0 1 .131-.232l9.642-9.642a.5.5 0 0 0-.642.056L6.854 4.854a.5.5 0 1 1-.708-.708L9.44.854A1.5 1.5 0 0 1 11.5.796a1.5 1.5 0 0 1 1.998-.001m-.644.766a.5.5 0 0 0-.707 0L1.95 11.756l-.764 3.057 3.057-.764L14.44 3.854a.5.5 0 0 0-.708-.708l-1.585-1.585z" />
                      </svg>
                    </button>
                    <button onClick={() => this.deleteData(person.ID)}>
                      <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" className="bi bi-trash3" viewBox="0 0 16 16">
                        <path d="M6.5 1h3a.5.5 0 0 1 .5.5v1H6v-1a.5.5 0 0 1 .5-.5M11 2.5v-1A1.5 1.5 0 0 0 9.5 0h-3A1.5 1.5 0 0 0 5 1.5v1H2.506a.58.58 0 0 0-.01 0H1.5a.5.5 0 0 0 0 1h.538l.853 10.66A2 2 0 0 0 4.885 16h6.23a2 2 0 0 0 1.994-1.84l.853-10.66h.538a.5.5 0 0 0 0-1h-.995a.59.59 0 0 0-.01 0zm1.958 1-.846 10.58a1 1 0 0 1-.997.92h-6.23a1 1 0 0 1-.997-.92L3.042 3.5zm-7.487 1a.5.5 0 0 1 .528.47l.5 8.5a.5.5 0 0 1-.998.06L5 5.03a.5.5 0 0 1 .47-.53Zm5.058 0a.5.5 0 0 1 .47.53l-.5 8.5a.5.5 0 1 1-.998-.06l.5-8.5a.5.5 0 0 1 .528-.47ZM8 4.5a.5.5 0 0 1 .5.5v8.5a.5.5 0 0 1-1 0V5a.5.5 0 0 1 .5-.5" />
                      </svg>
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
          <button className="btn btn-primary" onClick={() => this.onReturn()}>Return</button>
        </div>
      </section>
    );
  }
}
