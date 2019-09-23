import * as React from 'react';
import styles from './ApiUsingReactAdal.module.scss';
import { IApiUsingReactAdalProps } from './IApiUsingReactAdalProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { adalApiFetch, adalConfig } from '../../../common/adalConfig';
export interface IApiUsingReactAdalState {
  error: string;
  apiResponse: string;
  graphResponse: string;
}


export default class ApiUsingReactAdal extends React.Component<IApiUsingReactAdalProps, IApiUsingReactAdalState> {
  constructor(props: IApiUsingReactAdalProps, state: IApiUsingReactAdalState) {
    super(props);

    this.state = {
      error: '',
      apiResponse: '',
      graphResponse: ''
    };
  }
  public componentWillMount(): void {
    this._getCurrentUser();
    this._getGraphMe();
  }

  private _getCurrentUser() {
    try {
      //this.webapi.GetCampusWebAPI(APISource.TestAPI, "GetMe")
      adalApiFetch(adalConfig.clientId, "https://fn.azurewebsites.net/api/GetMe", null)
        .then(r => r.text())
        .then(r => {
          this.setState({
            apiResponse: r
          });
        });
    }
    catch (e) {
      console.error(e);
    }
  }
  private _getGraphMe() {
    try {
      adalApiFetch("https://graph.microsoft.com", "https://graph.microsoft.com/v1.0/me", null)
        .then(r => r.json())
        .then(r => {
          this.setState({
            graphResponse: r.displayName
          });
        });
    }
    catch (e) {
      console.error(e);
    }
  }

  public render(): React.ReactElement<IApiUsingReactAdalProps> {
    return (
      <div className={styles.apiUsingReactAdal}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>Welcome to All!</span>
              <p className={styles.description}>
                Graph output: {escape(this.state.graphResponse)}
                <br />
                API output: {escape(this.state.apiResponse)}
              </p>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
