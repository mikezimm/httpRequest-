
import { HttpClient, HttpClientResponse } from "@microsoft/sp-http";

const APISiteGetEndPoint : string = '_api/site';
const APISitePostQueryEndPoint : string = '_vti_bin/client.svc/ProcessQuery';

function buildGroupOwnerQueryBody( siteGuid: string, targetGroupId: string, ownerGroupID: string ) {

  let body: string = '';
  body += '<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="15.0.0.0" ApplicationName=".NET Library" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009">';
  body +=   '<SetProperty Id="1" ObjectPathId="2" Name="Owner">';
  body +=     '<Parameter ObjectPathId="3" />';
  body +=   '</SetProperty>';
  body +=   '<Method Name="Update" Id="4" ObjectPathId="2" />';
  body +=   '</Actions>';
  body +=   '<ObjectPaths>';
  body +=     '<Identity Id="2" Name="740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:AddSiteGUIDHERE:g:TargetGroupID" />';
  body +=     '<Identity Id="3" Name="740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:AddSiteGUIDHERE:g:OwnerGroupID" />';
  body +=   '</ObjectPaths>';
  body += '</Request>';

  body = body.replace(/AddSiteGUIDHERE/g, siteGuid );
  body = body.replace(/TargetGroupID/g, targetGroupId );
  body = body.replace(/OwnerGroupID/g, ownerGroupID );

  return body;

}

export async function functionUpdateGroup ( httpClient: HttpClient, siteUrl: string, siteGuid: string, targetGroupId: string, ownerGroupID: string  ) {

  if ( siteUrl.lastIndexOf('/') !== siteUrl.length -1 ) { siteUrl += '/';}

    const endpoint: string = `${ siteUrl }${ APISitePostQueryEndPoint }`;

    let body = buildGroupOwnerQueryBody( siteGuid, targetGroupId, ownerGroupID );

    const request: any = {
      body: body
    };

    let result = null;
    let errMessage = '';

    try {
      result = await httpClient.post( endpoint, HttpClient.configurations.v1, request);
    } catch (e) {
      console.log( 'httpERROR catch: ', e  );
    }

    console.log( result );
    
    return result;

}

export class UpdateGroup {

  constructor(private httpClient: HttpClient ) { }

  public updateOwner( siteUrl: string, siteGuid: string, targetGroupId: string, ownerGroupID: string ): Promise<any> {
    if ( siteUrl.lastIndexOf('/') !== siteUrl.length -1 ) { siteUrl += '/';}
    return new Promise<any>((resolve,reject) => {
      const endpoint: string = `${ siteUrl }${ APISitePostQueryEndPoint }`;

      let body = buildGroupOwnerQueryBody( siteGuid, targetGroupId, ownerGroupID );

      const request: any = {
        body: body
      };

      this.httpClient.post( endpoint, HttpClient.configurations.v1, request)
      .then((rawResponse: HttpClientResponse) => {
          return rawResponse.json();
      })
      .then((jsonResponse: any ) => {
          resolve(jsonResponse);
      })
      .catch(( error ) => {
        reject( error );
      });

    });
  }
}