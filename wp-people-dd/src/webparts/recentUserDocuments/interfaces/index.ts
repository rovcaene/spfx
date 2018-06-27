import { IUser } from "../../people/interfaces";

//used to map search result cells
export interface ICellValue {
  Key: string;
  Value: string;
}

//type that will store search results
export interface ISearchResult{
  url:string;
  id:string;
  title:string;
  previewImageUrl: string;
  lastModifiedByPhotoUrl?: string;
  lastModifiedByName?: string;
  lastModifiedTime?: string;
  extension?: string;
}

export interface IDataService{
  GetResults(user: IUser) : Promise<ISearchResult[]>;
}