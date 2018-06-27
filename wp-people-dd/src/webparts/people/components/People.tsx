import * as React from 'react';
import styles from './People.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { MSGraphClient } from '@microsoft/sp-client-preview';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { DataService } from '../dal/DataService';
import { IGraphMember, IDataService, IUser } from '../interfaces';
import { Spinner, SpinnerSize, FocusZone, Persona } from 'office-ui-fabric-react';
import * as _ from 'lodash';

export interface IPeopleProps {
  context: WebPartContext;
  selectedGroup?: number;
  onUserSelected?: (user: IUser) => void;
}

export interface IPeopleState {
  loading: boolean;
  users?: IUser[];
}

export default class People extends React.Component<IPeopleProps, IPeopleState> {
  private _dataService: IDataService;

  constructor(props) {
    super(props);
    this._dataService = new DataService(this.props.context);
    this.state = {
      loading: false,
      users: []
    };

    this.onPersonaSelected = this.onPersonaSelected.bind(this);
  }

  /**
   * Loads the users from the selected SharePoint Group.
   */
  private loadUsers(): void {
    if (this.props.selectedGroup) {
      this.setState({ loading: true });

      this._dataService.GetUsersFromSPGroup(this.props.selectedGroup)
        .then((users: IUser[]) => {
          const uniqUsers = _.unionWith(users, (el1, el2) => el1.id === el2.id);

          this.setState({ loading: false, users: uniqUsers });
          users.forEach(u => {
            this._dataService.GetUserImage(u.id)
              .then((s: string) => {
                u.photo = s;
                this.setState({ users: uniqUsers });
              });
          });
        });
    }
  }

  /**
   * This method is invoked before a mounted component receives new props. 
   * @param nextProps The new props.
   */
  public componentWillReceiveProps(nextProps: Readonly<IPeopleProps>, nextContext: any): void {
    if (this.props.selectedGroup !== nextProps.selectedGroup) {
      this.setState({ users: [] });
    }
  }

  /**
   * This method is invoked immediately after updating occurs. This method is not called for the initial render.
   * @param prevProps Previous properties. 
   * @param prevState Previous state.
   * @param prevContext Previous context.
   */
  public componentDidUpdate(prevProps: Readonly<IPeopleProps>, prevState: Readonly<IPeopleState>, prevContext: any): void {
    if (this.props.selectedGroup !== prevProps.selectedGroup) {
      this.loadUsers();
    }
  }

  /**
   * This method is invoked immediately after a component is mounted.
   */
  public componentDidMount(): void {
    this.loadUsers();
  }

  protected onPersonaSelected(user: IUser){
    if(this.props.onUserSelected)
      this.props.onUserSelected(user);
  }

  public render(): React.ReactElement<IPeopleProps> {
    return (
      <div className={styles.people}>
        <div className={styles.container}>
          {
            (this.state.loading) ?
              <Spinner size={SpinnerSize.large} /> :
              (this.state.users.length === 0) ?
                <div className={styles.description}>No members found ...</div> :
                this.state.users.map((user: IUser) => {
                  return <div className={styles.card}>
                    <FocusZone>
                      <Persona key={user.id} primaryText={user.name} imageUrl={user.photo} className={styles.persona} onClick={() => this.onPersonaSelected(user)} />
                    </FocusZone>
                  </div>;
                })
          }
        </div>
      </div>
    );
  }
}
