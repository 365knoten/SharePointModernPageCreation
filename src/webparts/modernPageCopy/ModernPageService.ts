import { Log, ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { PageContext } from "@microsoft/sp-page-context";
import { SPHttpClient } from "@microsoft/sp-http";

const LOG_SOURCE: string = "ModernPageService:IModernPageService";

export interface IModernPageServiceConfiguration {
    listID: string,
    itemID: number,
}

export interface IBasicPage {
    CanvasContent1: string;
    ID: number;
}


export interface IModernPageService {
    copyPage: (templateSiteRelativeUrl: string, sourceName: string, destinationName: string) => Promise<void>
}

/**
* A service that encapsulates all calls to SharePoint as a reusable Service
*/
export class ModernPageService implements IModernPageService {


    private _httpClient: SPHttpClient;
    private _currentWebUrl: string;
    private _currentWebServerRelativePath: string;
    private _hostname: string;


    //Create a ServiceKey which will be used to consume the service.
    public static readonly serviceKey: ServiceKey<IModernPageService> =
        ServiceKey.create<ModernPageService>(LOG_SOURCE, ModernPageService);

    constructor(serviceScope: ServiceScope) {
        serviceScope.whenFinished(() => {

            // Consume the pagecontext from dependency injection 
            const pageContext = serviceScope.consume(PageContext.serviceKey);
            this._httpClient = serviceScope.consume(SPHttpClient.serviceKey);

            //construct some variables
            this._currentWebUrl = pageContext.web.absoluteUrl;
            const url = new URL(this._currentWebUrl);
            this._currentWebServerRelativePath = url.pathname;
            this._hostname=`https://${url.hostname}`


            Log.info(
                LOG_SOURCE,
                "Constructed ModernPageService"
            );
        });
    }

    public async copyPage(templateSiteRelativeUrl: string, sourceName: string, destinationName: string) {
        Log.info(
            LOG_SOURCE,
            `Copy Modern Page ${sourceName} from web ${templateSiteRelativeUrl} as page ${destinationName}`
        );

        const baseSite = await this.getPageByName(sourceName, `${this._hostname}${templateSiteRelativeUrl}`);
        await this.createPage(destinationName);
        const targetPage = await this.getPageByName(destinationName);
        if (targetPage !== null && baseSite !== null) {
            await this.updateCanvasContent(this._currentWebUrl, targetPage?.ID, baseSite?.CanvasContent1)
            console.log("Done")
        }
    }





    private async getPageByName(pagename: string, weburl?: string) {
        Log.info(
            LOG_SOURCE,
            `Getting Page by name ${pagename} from web ${weburl !== undefined ? weburl : this._currentWebUrl}`
        );


        const APICALL = `${weburl !== undefined ? weburl : this._currentWebUrl}/_api/web/lists/GetByTitle('Site Pages')/items??$select=ID,Title,CanvasContent1&$filter=FileLeafRef eq '${pagename}.aspx'`

        const page = await this._httpClient.get(APICALL, SPHttpClient.configurations.v1)
            .then((response) => { return response.json() });

        if (page !== null && page.value !== null && page.value.length > 0) {
            return page.value[0] as IBasicPage
        }
        return null;
    }


    private async createPage(pagename: string) {
        Log.info(
            LOG_SOURCE,
            `Creating Page with name ${pagename} in web ${this._currentWebUrl}`
        );

        const APICALL = `${this._currentWebUrl}/_api/Web/GetFileByServerRelativeUrl('${this._currentWebServerRelativePath}/SitePages/Home.aspx')/CopyTo(strnewurl='${this._currentWebServerRelativePath}/SitePages/${pagename}.aspx',bOverwrite=true)`

        const spOpts = {
            body: ``
        };
        return await this._httpClient.post(APICALL, SPHttpClient.configurations.v1, spOpts)
            .then((response) => { return response.text() }) as IBasicPage

    }


    private async updateCanvasContent(templateSiteUrl: string, pageID: number, content: string) {
        Log.info(
            LOG_SOURCE,
            `Updating Content of Page with ID ${pageID} to content "${content}"`
        );

        const APICALL = `${this._currentWebUrl}/_api/lists/getbytitle('Site Pages')/items(${pageID})`

        const headers = {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=verbose',
            'odata-version': '',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE'
        };

        const body: string = JSON.stringify({
            __metadata: { type: "SP.Data.SitePagesItem" },
            CanvasContent1: content
        });

        const options = {
            body: body,
            headers: headers,
            method: 'POST'
        };

        return await this._httpClient.post(APICALL, SPHttpClient.configurations.v1, options);

    }


}