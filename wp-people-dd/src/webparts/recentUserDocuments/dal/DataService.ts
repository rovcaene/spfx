import { IDataService, ISearchResult, ICellValue } from "../interfaces";
import { IWebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient } from "@microsoft/sp-http";
import { IUser } from "../../people/interfaces";

export default class DataService implements IDataService {
  constructor(private context: IWebPartContext){

  }

  private getResultUrl(result: ICellValue[]): string{
    var itemUrl = this.getValueFromResults('ServerRedirectedURL', result);
    itemUrl = itemUrl? itemUrl: this.getValueFromResults('OriginalPath', result);
    return itemUrl;
  }

   private getPreviewImageUrl(result: ICellValue[], siteUrl: string): string {
    const uniqueID: string = this.getValueFromResults('uniqueID', result);
    const siteId: string = this.getValueFromResults('siteID', result);
    const webId: string = this.getValueFromResults('webID', result);
    const docId: string = this.getValueFromResults('DocId', result);
    if (uniqueID !== null && siteId !== null && webId !== null && docId !== null) {
      return `${siteUrl}/_layouts/15/getpreview.ashx?guidFile=${uniqueID}&guidSite=${siteId}&guidWeb=${webId}&docid=${docId}
      &metadatatoken=300x424x2&ClientType=CodenameOsloWeb&size=small`;
    }
    else {
      return '';
    }
  }

  private getValueFromResults(key: string, results: ICellValue[]): string {
    let value: string = '';

    if (results != null && results.length > 0 && key != null) {
      for (let i: number = 0; i < results.length; i++) {
        const resultItem: ICellValue = results[i];
        if (resultItem.Key === key) {
          value = resultItem.Value;
          break;
        }
      }
    }

    return value;
  }

  private trim(s: string): string {
    if (s != null && s.length > 0) {
      return s.replace(/^\s+|\s+$/gm, '');
    }
    else {
      return s;
    }
  }

  private getUserPhotoUrl(userEmail: string, siteUrl: string): string {
    //return '';
    return `${siteUrl}/_layouts/15/userphoto.aspx?size=S&accountname=${userEmail}`;
  }

  public GetResults(user: IUser) : Promise<ISearchResult[]>{
    // Check if the query field changed
    let query = 'fileType:docx';
    if(user){
      query += ` and author:${user.name}`;
    }
    const siteUrl = this.context.pageContext.web.absoluteUrl;
    const searchUrl = `${this.context.pageContext.web.absoluteUrl}/_api/search/query?querytext='${query}'&selectproperties='Title,Author,FileExtension,DocId, FileType,EditorOwsUser,LastModifiedTime,uniqueID,webID,siteID'&clienttype='ContentSearchRegular'`;
    
    return this.context.spHttpClient
      .get(searchUrl, SPHttpClient.configurations.v1, { headers: { 'odata-version': '3.0' } })
      .then(response => { return response.json(); })
      .then(data => {
        const searchContent: ISearchResult[] = [];
          if (data.PrimaryQueryResult &&
              data.PrimaryQueryResult.RelevantResults &&
              data.PrimaryQueryResult.RelevantResults.Table.Rows.length > 0) {
                console.log("total items found:" + data.PrimaryQueryResult.RelevantResults.Table.Rows.length);
                data.PrimaryQueryResult.RelevantResults.Table.Rows.forEach((row: any): void => {
                  const cells: ICellValue[] = row.Cells;
                  var editorInfo: string[];
                  if(this.getValueFromResults('EditorOwsUser', cells)!==null) editorInfo = this.getValueFromResults('EditorOwsUser', cells).split('|');
                  const modifiedDate: Date = this.getValueFromResults('LastModifiedTime', cells) === null ? new Date(): new Date(this.getValueFromResults('LastModifiedTime', cells).replace('.0000000', ''));
                  const dateString: string = (modifiedDate.getMonth() + 1) + '/' + modifiedDate.getDate() + '/' + modifiedDate.getFullYear();
                  console.log('Adding search result:' + this.getValueFromResults('Title', cells));
                  console.log('Extension:' + this.getValueFromResults('FileExtension', cells));
                  console.log('Type:' + this.getValueFromResults('FileType', cells));
                  searchContent.push({
                    id: this.getValueFromResults('DocId', cells),
                    url: this.getResultUrl(cells),
                    title: this.getValueFromResults('Title', cells),
                    previewImageUrl: this.getPreviewImageUrl(cells, siteUrl),
                    lastModifiedTime: dateString,
                    lastModifiedByName: editorInfo ? this.trim(editorInfo[1]):"",
                    lastModifiedByPhotoUrl: editorInfo ? this.getUserPhotoUrl(this.trim(editorInfo[0]), siteUrl):"",
                    extension: this.getValueFromResults('FileType', cells)
                  });
                  console.log('Added search result:' + this.getValueFromResults('Title', cells));

                });
                return searchContent;
          }
      });
  }
}