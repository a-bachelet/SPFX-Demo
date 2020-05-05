import * as React from 'react';
import styles from './FirstWebpart.module.scss';
import { IFirstWebpartProps } from './IFirstWebpartProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IFirstWebpartState } from './IFirstWebpartState';
import { UserService } from '../services/UserService';
import { User } from '@microsoft/microsoft-graph-types';

export default class FirstWebpart extends React.Component<IFirstWebpartProps, IFirstWebpartState> {
  private _userService: UserService;

  public render(): React.ReactElement<IFirstWebpartProps> {
    return (
      <div className={ styles.firstWebpart }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              { !this.state.user ?
                <p>Chargement en cours...</p> :
                <h1>{this.state.user.displayName}</h1>
              }
            </div>
          </div>
        </div>
      </div>
    );
  }

  public componentWillMount(): void {
    this.setState({
      user: null
    });

    this._userService = UserService.getInstance(this.props.msGraphClient);
  }

  public componentDidMount(): void {
    this._userService.getProfile().then((user: User) => {
      this.setState({
        user
      });
    });
  }
}
