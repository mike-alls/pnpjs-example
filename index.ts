//import { sp } from "@pnp/sp";
import sp = require("@pnp/sp");
import "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/folders";
//import { IFolders, Folders } from "@pnp/sp/folders";
import "@pnp/sp/files";
//import { IFiles, Files } from "@pnp/sp/files";
import "@pnp/sp/search";
//import { ISearchQuery, SearchResults, SearchQueryBuilder } from "@pnp/sp/search";
import "@pnp/sp/site-users/web";

//import { default as Log } from 'microservice-base/app/lib/log';
//import config from 'microservice-base/app/config';

//const log = Log.createLog({ name: __filename, level: 'info' });

//const SHAREPOINT_ACCESS_TOKEN = config.secrets.sharepoint.oauth;
//const SHAREPOINT_USER_ID = config.secrets.sharepoint.userid;

//export const PREVIEW_FILE_TYPES_PDF = '.ai, .doc, .docm, .docx, .eps, .gdoc, .gslides, .odp, .odt, .pps, .ppsm, .ppsx, .ppt, .pptm, .pptx, .rtf'.split(', ');
//export const PREVIEW_FILE_TYPES_HTML = '.csv, .ods, .xls, .xlsm, .gsheet, .xlsx'.split(', ');

import { SPFetchClient } from "@pnp/nodejs-commonjs";
//import { sp } from "@pnp/sp-commonjs";

sp.setup({
    sp: {
        fetchClientFactory: () => {
            return new SPFetchClient("{ site url }", "{ client id }", "{ client secret }");
        },
    },
});

/*     sp.setup({
      spfxContext: this.context
    });
 */
class SharePointSDK {
// Use PnPJS Profiles of SP.Utilities.Utiliy searchPrincipals to search for users or groups
// use "const users = await sp.web.siteUsers();" to get all users of a site
  public search(query: string, options: {
    maxResults?: number,
    justFiles?: boolean,
  }, cbSearch: {
    (error?: any, response?: any): void;
  }) {
    const msgHdr = 'search: ';

/*    sp.setup({
      spfxContext: this.context
    }); */

	let maxResults = 20;
    if (options.maxResults) {
      maxResults = options.maxResults;
    }

    let justFiles = true;
    if (options.justFiles) {
      justFiles = options.justFiles;
    }

    const fileCategories: any = ['document', 'pdf', 'spreadsheet', 'presentation', 'image'];
    
/*	const results: SearchResults = await sp.search(<ISearchQuery>{
		Querytext: query,
		RowLimit: maxResults,
		EnableInterleaving: true,
	}); */

	const spQueryParams: any = {
		Querytext: query,
		RowLimit: maxResults,
		EnableInterleaving: true,
    };

//    log.debug(msgHdr + 'spQueryParams = ', spQueryParams);
//    sp.search(<ISearchQuery> spQueryParams)
    sp.search(spQueryParams)
      .then((results) => {
//        log.debug(msgHdr + 'search.result = ', JSON.stringify(results.PrimarySearchResults, null, 4));
        return cbSearch(null, results.PrimarySearchResults);
      })
      .catch((err) => {
//        log.error(msgHdr + 'err = ', err);
        return cbSearch(err);
      });
  }

  public getFiles(path: string, cbGetFiles: {
    (error?: any, response?: any): void;
  }) {
    const msgHdr = 'getFiles: ';

/*    sp.setup({
      spfxContext: this.context
    }); */

	sp.web.getFolderByServerRelativePath(path).files()
      .then((foundFiles) => {
//        log.debug(msgHdr + 'foundFiles.result = ', JSON.stringify(foundFiles.result, null, 4));
/*        const items = foundFiles;
        items.entries = items.entries.sort((a, b) => {
          return a.name.localeCompare(b.name, undefined, { numeric: true });
        }); */
        return cbGetFiles(null, foundFiles); //Shouldn't this return items instead of foundFiles?
      })
      .catch((err) => {
//        log.error(msgHdr + 'err = ', err);
        return cbGetFiles(err);
      });
  }

  public getFolders(path: string, cbGetFolders: {
    (error?: any, response?: any): void;
  }) {
    const msgHdr = 'getFolders: ';

/*    sp.setup({
      spfxContext: this.context
    }); */

	sp.web.getFolderByServerRelativePath(path).folders()
      .then((foundFolders) => {
//        log.debug(msgHdr + 'foundFolders.result = ', JSON.stringify(foundFolders.result, null, 4));
/*        const items = foundFolders;
        items.entries = items.entries.sort((a, b) => {
          return a.name.localeCompare(b.name, undefined, { numeric: true });
        }); */
        return cbGetFolders(null, foundFolders); //Shouldn't this return items instead of foundFiles?
      })
      .catch((err) => {
//        log.error(msgHdr + 'err = ', err);
        return cbGetFolders(err);
      });
  }


  public getFileContent(path: string, cbGetFileContent: {
    (error?: any, response?: any): void;
  }) {
    const msgHdr = 'getFileContent: ';

/*    sp.setup({
      spfxContext: this.context
    }); */

	sp.web.getFileByServerRelativeUrl(path).getBlob()
      .then((blob) => {
        return cbGetFileContent(null, blob);
      })
      .catch((err) => {
//        log.error(msgHdr + 'err = ', err);
        return cbGetFileContent(err);
      });
  }
}

export default new SharePointSDK();
