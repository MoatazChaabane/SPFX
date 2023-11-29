import * as React from 'react';
import styles from './Sol1.module.scss';
// import type { ISol1Props } from './ISol1Props';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface ISol1Props {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext; // Add this line to include the 'context' property
}

export default class Sol1 extends React.Component<ISol1Props, {}> {

  private sendData(): void {
    //les donées des inputs
    
  
    const url = `${this.props.context.pageContext.web.absoluteUrl}/sites/ABC/_api/web/lists/getbytitle('personne')/items`;
   
   
//GET Api
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
    response.json().then((responseJSON: JSON) => {
      console.log(responseJSON);
    });
  } else {
    console.error('Error:', response.statusText);
  }
})
.catch((error) => {
  console.error('Error:', error);
});
  }

  public render(): React.ReactElement<ISol1Props> {
    const {
      
      hasTeamsContext,
    } = this.props;

    return (
      <section className={`${styles.sol1} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
       <h1>TEST</h1>
       </div>
       <button onClick={this.sendData}></button>
      </section>
    );
  }
}
