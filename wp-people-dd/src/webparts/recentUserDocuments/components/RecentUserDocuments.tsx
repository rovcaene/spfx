import * as React from 'react';
import styles from './RecentUserDocuments.module.scss';
import { IUser } from '../../people/interfaces';
import { IWebPartContext } from '@microsoft/sp-webpart-base';
import { IDataService, ISearchResult } from '../interfaces';
import DataService from '../dal/DataService';
import { DocumentCard, DocumentCardPreview, DocumentCardTitle, DocumentCardActivity } from 'office-ui-fabric-react';

export interface IRecentUserDocumentsProps {
  context: IWebPartContext;
  user?: IUser;
}

export interface IRecentUserDocumentsState {
  results: ISearchResult[];
}

export default class RecentUserDocuments extends React.Component<IRecentUserDocumentsProps, IRecentUserDocumentsState> {
  private _dataService: IDataService;

  constructor(props) {
    super(props);
    this._dataService = new DataService(this.props.context);
    this.state = { results: [] };
  }

  /**
   * This method is invoked immediately after updating occurs. This method is not called for the initial render.
   * @param prevProps Previous properties. 
   * @param prevState Previous state.
   * @param prevContext Previous context.
   */
  public componentDidUpdate(prevProps: Readonly<IRecentUserDocumentsProps>, prevState: Readonly<IRecentUserDocumentsState>, prevContext: any): void {
    if (this.props.user !== prevProps.user) {
      this.loadResults();
    }
  }

  /**
   * This method is invoked immediately after a component is mounted.
   */
  public componentDidMount(): void {
    this.loadResults();
  }

  private loadResults(): void {
    this._dataService
      .GetResults(this.props.user)
      .then(results => this.setState({ results }));
  }

  public render(): React.ReactElement<IRecentUserDocumentsProps> {
    return (
      <div className={styles.recentUserDocuments}>
        <div className={styles.container}>
          {
            this.state.results.map((resultItem) => {
              const iconUrl: string = `https://spoprod-a.akamaihd.net/files/odsp-next-prod_ship-2016-08-15_20160815.002/odsp-media/images/filetypes/32/${resultItem.extension}.png`;
              return (
                <DocumentCard onClickHref={resultItem.url} key={resultItem.id}>
                  <DocumentCardPreview
                    previewImages={[
                      {
                        previewImageSrc: resultItem.previewImageUrl,
                        iconSrc: iconUrl,
                        width: 294,
                        height: 196,
                        accentColor: '#ce4b1f'
                      }
                    ]}
                  />
                  <DocumentCardTitle title={resultItem.title} />
                  <DocumentCardActivity
                    activity={'Modified ' + resultItem.lastModifiedTime}
                    people={
                      [
                        { name: resultItem.lastModifiedByName, profileImageSrc: resultItem.lastModifiedByPhotoUrl }
                      ]
                    }
                  />
                </DocumentCard>
              );
            })
          }
        </div>
      </div>
    );
  }
}
