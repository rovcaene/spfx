import { IDataService, ISPMemberInfo, MemberType, ISPUser, IGraphMember, IUser } from '../interfaces/index';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse, HttpClientResponse } from '@microsoft/sp-http';
import { MSGraphClient } from '@microsoft/sp-client-preview';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

/** Data Service for the retrieval of all web service related data. */
export class DataService implements IDataService {
  private msGraphClient: MSGraphClient;

  constructor(private context: WebPartContext) {
    this.msGraphClient = context.serviceScope.consume(MSGraphClient.serviceKey);
  }

  /**
   * Retrieves the memberinfo from the SharePoint site.
   * @returns Promise that will return an array of site members once resolved.
   */
  public GetMemberInfo(): Promise<ISPMemberInfo[]> {
    return this.context.spHttpClient
      .get(`${this.context.pageContext.web.absoluteUrl}/_api/Web/roleassignments?$expand=member`, SPHttpClient.configurations.v1)
      .then(data => data.json())
      .then(jsonData => {
        return jsonData.value.map(roleAssignment => {
          return {
            PrincipalId: roleAssignment.PrincipalId,
            Id: roleAssignment.Member.Id,
            Title: roleAssignment.Member.Title,
            Type: roleAssignment.Member.PrincipalType === 8 ? MemberType.Group : MemberType.User
          } as ISPMemberInfo;
        });
      });
  }

  /** 
   * Retrieves the users, including additional information, contained within a SharePoint Group. Active Directory groups will be resolved to include the users contained within.
   * @param spGroupId The ID of the SharePoint Group
   * @returns Promise that will return an array of all the users.
  */
  public GetUsersFromSPGroup(spGroupId: number): Promise<IUser[]> {
    return this.context.spHttpClient
      .get(`${this.context.pageContext.web.absoluteUrl}/_api/Web/sitegroups(${spGroupId})/Users`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse): Promise<{ value: ISPUser[] }> => {
        return response.json();
      })
      .then((data: { value: ISPUser[] }) => {
        let promises: Promise<IGraphMember[]>[] = [];

        data.value.forEach((element: ISPUser) => {
          if (element.PrincipalType === 1) {
            // USER
            let userId: string = element.LoginName.substr(element.LoginName.lastIndexOf('|') + 1);
            console.log(`UserID: ${userId}`);
            promises.push(Promise.resolve<IGraphMember[]>([{ id: userId, displayName: element.Title }]));
          }
          else if (element.PrincipalType === 4) {
            // GROUP
            let groupId: string = element.LoginName.substr(element.LoginName.lastIndexOf('|') + 1);
            console.log(`GroupID: ${groupId}`);
            promises.push(this.GetUsersFromADGroup(groupId));
          }
        });

        return Promise.all(promises);
      })
      .then((data: IGraphMember[][]) => {
        // Flatten the array
        let gUsers: IGraphMember[] = [].concat.apply([], data);

        // Return the users
        return gUsers.map((gUser: IGraphMember) => {
          return { id: gUser.id, name: gUser.displayName };
        });
      });
  }

  /**
   * Retrieves all members contained within an Active Directory group. Remark: groups are not recursivey resolved!
   * @param groupId The ID of the Active Directory Group
   * @returns Promise that will return an array of members contained within the group.
   */
  private GetUsersFromADGroup(groupId: string): Promise<IGraphMember[]> {
    return this.msGraphClient
      .api(`/groups/${groupId}/members`)
      .get()
      .then((members: { value: MicrosoftGraph.User[] }) => {
        return members.value.map(member => ({ id: member.mail, displayName: member.displayName }));
      });
  }

  /**
   * Retrieves the image for the supplied user.
   * @param userId The ID of the user.
   * @returns Promise containing the blob url string for the image.
   */
  public GetUserImage(userId: string): Promise<string> {
    return this.GetImage(userId, 'users');
  }

  /**
   * Retrieves the image for the supplied AD group.
   * @param groupId The ID of the AD group.
   * @returns Promise containing the blob url string for the image.
   */
  public GetGroupImage(groupId: string): Promise<string> {
    return this.GetImage(groupId, 'groups');
  }

  public GetImage(id: string, type: string): Promise<string> {
    return this.msGraphClient
      .api(`/${type}/${id}/photos/48x48/$value`, { defaultVersion: 'beta' })
      .responseType('blob')
      .get()
      .then((blob: Blob) => {
        return blob ? URL.createObjectURL(blob) : null;
      });
  }
}