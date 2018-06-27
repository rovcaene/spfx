export interface ISPMemberInfo{
  PrincipalId:number;
  Id:number;
  Title:string;
  Type:MemberType;
}

export enum MemberType{
  Group,
  User
}

export interface ISPUser{
  Id: number;
  PrincipalType: number;
  LoginName: string;
  Title: string;
}

export interface IGraphMember{
  id: string;
  displayName: string;
  photo?: string;
}

export interface IUser{
  id: string;
  name: string;
  photo?: string;
}

export interface IDataService{
  GetMemberInfo() : Promise<ISPMemberInfo[]>;
  GetUsersFromSPGroup(groupId: number) : Promise<IUser[]>;
  GetUserImage(userId: string) : Promise<string>;
  GetGroupImage(groupId: string) : Promise<string>;
}