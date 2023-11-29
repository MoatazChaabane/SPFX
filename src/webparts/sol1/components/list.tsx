import * as React from 'react';
import styles from './Sol1.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

interface IPerson {
  Title: string;
  Email: string;
  Age: number;
}

export interface IListProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
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

  componentDidMount() {
    this.getData();
  }

  private getData(): void {
    const url = `https://mch12.sharepoint.com/sites/ABC/_api/web/lists/getbytitle('personne')/items`;

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
              </tr>
            </thead>
            <tbody>
              {people.map((person, index) => (
                <tr key={index}>
                  <td>{person.Title}</td>
                  <td>{person.Email}</td>
                  <td>{person.Age}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>
    );
  }
}
